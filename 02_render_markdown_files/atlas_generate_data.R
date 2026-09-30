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

commitments_2026_path <- here(
  "01_projects_repository",
  "2026_water_data_challenge",
  "2026 CA Open Water Data Showcase _ Project Commitment (Responses).xlsx"
)
if (!file.exists(commitments_2026_path)) {
  stop("2026 commitments workbook not found: ", commitments_2026_path)
}

projects_raw <- read_excel(
  path = source_path,
  sheet = "projects_table",
  col_types = "text"
) %>%
  clean_names() %>%
  filter(publish_to_website == TRUE)

commitments_2026_raw <- read_excel(
  path = commitments_2026_path,
  sheet = "Form Responses 1",
  col_types = "text"
) %>%
  clean_names()

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
    status = "Historical",
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
    title,
    year,
    event,
    status,
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

extract_team_name <- function(x) {
  x <- replace_na(x, "")
  m <- str_match(x, regex("team\\s*name\\s*:\\s*([^\\r\\n]+)", ignore_case = TRUE))
  out <- ifelse(!is.na(m[, 2]), str_trim(m[, 2]), "")
  replace_na(out, "")
}

extract_year_from_timestamp <- function(x) {
  x <- replace_na(as.character(x), "")

  year_prefix <- suppressWarnings(as.numeric(str_match(x, "^((?:19|20)\\d{2})")[, 2]))

  serial_num <- suppressWarnings(as.numeric(x))
  serial_year <- suppressWarnings(as.numeric(format(as.Date(serial_num, origin = "1899-12-30"), "%Y")))
  serial_year <- ifelse(!is.na(serial_num) & serial_num > 20000 & serial_num < 80000, serial_year, NA_real_)

  year <- dplyr::coalesce(year_prefix, serial_year)
  ifelse(is.na(year), 2026, year)
}

commitments_2026 <- commitments_2026_raw %>%
  mutate(
    year = extract_year_from_timestamp(timestamp),
    title = replace_na(what_is_your_project_name, ""),
    description = replace_na(
      briefly_describe_the_open_water_data_project_or_opportunity_you_are_hoping_to_explore_please_do_so_in_200_words_or_fewer,
      ""
    ),
    team_source = replace_na(
      if_this_is_a_team_project_please_provide_the_other_team_members_names_emails_and_the_team_name_if_you_have_one,
      ""
    ),
    submitter = str_trim(paste0(
      replace_na(your_name, ""),
      ifelse(replace_na(you_email_address, "") != "", paste0(" <", you_email_address, ">"), "")
    )),
    team_name = extract_team_name(team_source),
    team_members = str_trim(paste(
      ifelse(submitter != "", paste0("Submitter: ", submitter), ""),
      team_source,
      sep = "\n"
    )),
    team_members = str_replace(team_members, "^[\\n]+|[\\n]+$", ""),
    topics = "",
    topic_list = replicate(n(), character(0), simplify = FALSE),
    event = "2026 CA Open Water Data Showcase (Project Commitments)",
    raw_text = paste(
      title,
      description,
      replace_na(please_check_any_and_all_boxes_that_apply_i_am, ""),
      replace_na(
        please_list_any_key_partners_stakeholders_and_or_communities_that_engage_or_intersect_with_this_project_or_its_data_i_e_are_affected_by_the_outcomes_associated_with_this_data_opportunity,
        ""
      )
    ),
    keywords = lapply(raw_text, extract_keywords),
    project_url = NA_character_,
    status = "In progress",
    spatial_status = "statewide_no_geospatial_footprint",
    spatial_category = "California statewide",
    spatial_notes = "No reliable project-level geometry was recorded in the commitments workbook. This project is represented at the California statewide level until a documented footprint is added.",
    latitude = 36.7783,
    longitude = -119.4179
  ) %>%
  select(
    title,
    year,
    event,
    status,
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

projects <- bind_rows(projects, commitments_2026) %>%
  mutate(id = row_number(), .before = title)

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
