if (!requireNamespace("pacman", quietly = TRUE)) {
  install.packages("pacman", repos = "https://cloud.r-project.org")
}

pacman::p_load(
  here,
  readxl,
  dplyr,
  tidyr,
  stringr,
  jsonlite,
  janitor
)

source_path <- here("01_projects_repository", "project_repository_tables.xlsm")
if (!file.exists(source_path)) {
  stop("Project workbook not found: ", source_path)
}

projects_raw <- read_excel(
  path = source_path,
  sheet = "projects_table",
  col_types = "text"
) %>%
  clean_names() %>%
  filter(publish_to_website == TRUE)

keyword_keep <- c(
  "water", "data", "groundwater", "drought", "quality", "infrastructure",
  "monitoring", "mapping", "watershed", "flood", "flooding", "stormwater",
  "supply", "demand", "recharge", "reservoir", "salinity", "treatment",
  "reuse", "planning", "climate", "risk", "equity", "affordability",
  "ecosystem", "environmental", "model", "models", "modeling", "dashboard",
  "analytics", "sensor", "streamflow", "aquifer", "basin", "wetland", "wetlands",
  "contaminant", "contaminants", "drinking", "well", "wells", "rainfall",
  "runoff", "storage", "agriculture", "urban", "rural", "wateruse", "waterquality",
  "monitor", "forecast", "watersheds", "observations", "irrigation"
)

keyword_stopwords <- c(
  "about", "above", "across", "after", "again", "against", "along", "also",
  "among", "an", "and", "another", "any", "around", "before", "being",
  "below", "between", "both", "but", "by", "came", "can", "could", "day",
  "did", "does", "doing", "down", "each", "few", "find", "for", "from",
  "further", "get", "gets", "given", "had", "has", "have", "having", "here",
  "however", "into", "its", "just", "like", "made", "make", "many", "more",
  "most", "much", "need", "needed", "new", "not", "off", "often", "old",
  "once", "other", "our", "out", "over", "own", "part", "people", "same",
  "see", "show", "site", "sites", "some", "such", "take", "than", "that",
  "their", "them", "then", "there", "these", "they", "this", "those",
  "through", "under", "until", "upon", "use", "used", "using", "very", "want",
  "was", "way", "we", "were", "what", "when", "where", "which", "while",
  "who", "why", "will", "with", "within", "without", "work", "working",
  "would", "year", "years", "you", "your", "challenge", "california", "state",
  "county", "system", "systems", "application", "applications", "user", "users",
  "platform", "dataset", "datasets", "analysis", "analyses", "approach", "results",
  "result", "page", "pages", "website", "view", "views", "build", "built"
)

keyword_noise <- c(
  "project", "projects", "team", "teams", "support", "supports", "supporting",
  "time", "times", "tool", "tools", "resource", "resources", "irain",
  "information", "community", "digital", "online", "visual", "interface",
  "innovation", "story", "stories", "improve", "improving", "solution", "solutions",
  "research", "develop", "developing", "program", "programs", "service", "services",
  "decision", "framework", "impact", "impacts", "future", "current", "historical",
  "local", "regional", "statewide", "broader", "better", "value", "benefit", "benefits",
  "public", "private", "database", "datahub", "portal", "website"
)

extract_keywords <- function(text) {
  text <- ifelse(is.na(text), "", text)
  text <- str_replace_all(text, "https?://\\S+|www\\.\\S+", " ")
  text <- str_replace_all(text, "[^A-Za-z0-9\\s]", " ")
  text <- str_to_lower(text)

  tokens <- unlist(str_split(text, "\\s+"))
  tokens <- tokens[nchar(tokens) > 2]
  tokens <- tokens[!tokens %in% keyword_stopwords]
  tokens <- tokens[!tokens %in% keyword_noise]
  tokens <- tokens[tokens %in% keyword_keep | (nchar(tokens) > 5 & !tokens %in% keyword_stopwords)]
  tokens <- unique(tokens)
  tokens[order(tokens)]
}

projects <- projects_raw %>%
  mutate(
    project_id = row_number(),
    .before = year,
    year = suppressWarnings(as.numeric(str_replace(as.character(year), "\\.0$", ""))),
    title = replace_na(title, ""),
    description = replace_na(description, ""),
    team_name = replace_na(team_name, ""),
    team_members = replace_na(team_members, ""),
    topics = replace_na(topics, ""),
    event = replace_na(event, ""),
    project_slug = title %>%
      str_to_lower() %>%
      make_clean_names() %>%
      str_sub(1, 50),
    project_url = paste0("projects/", project_slug, ".html"),
    raw_text = paste(title, description),
    keywords = lapply(raw_text, extract_keywords),
    spatial_status = "statewide_no_geospatial_footprint",
    spatial_category = "California statewide",
    spatial_notes = "No reliable project-level geometry was recorded in the workbook. This project is represented at the California statewide level until a documented footprint is added.",
    latitude = 36.7783,
    longitude = -119.4179
  ) %>%
  mutate(
    topic_list = str_split(topics, "\\|") %>%
      lapply(function(x) {
        x %>%
          str_replace_all(fixed('"'), "") %>%
          str_replace_all(fixed("'"), "") %>%
          str_replace_all(fixed("Topic: "), "") %>%
          str_trim() %>%
          .[. != "" & . != "NA"]
      })
  ) %>%
  select(
    id = project_id,
    title,
    year,
    event,
    team_name,
    team_members,
    description,
    topics,
    topic_list,
    keywords,
    project_url,
    spatial_status,
    spatial_category,
    spatial_notes,
    latitude,
    longitude
  )

output_dir <- here("docs")
if (!dir.exists(output_dir)) {
  dir.create(output_dir, recursive = TRUE)
}

write_json(
  x = projects,
  path = file.path(output_dir, "atlas-projects.json"),
  pretty = TRUE,
  digits = 8,
  auto_unbox = TRUE,
  null = "null"
)

message("Atlas dataset written to: ", file.path(output_dir, "atlas-projects.json"))
