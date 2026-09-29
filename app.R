# app.R
# Data center capacity dashboard
# Run with: shiny::runApp() from this folder (after running data_prep.R once).

library(shiny)
library(bslib)
library(leaflet)
library(leaflet.extras)
library(DT)
library(dplyr)
library(stringr)
library(DBI)
library(RSQLite)
library(readr)

DB_PATH <- "dc_data.sqlite"

if (!file.exists(DB_PATH)) {
  stop("dc_data.sqlite not found. Run `Rscript data_prep.R your_file.csv` first.")
}

# ---------- DB helpers ----------

get_con <- function() dbConnect(RSQLite::SQLite(), DB_PATH)

load_current <- function() {
  con <- get_con(); on.exit(dbDisconnect(con))
  dbReadTable(con, "dc_current") %>%
    as_tibble() %>%
    mutate(is_us = as.logical(is_us))
}

list_versions <- function() {
  con <- get_con(); on.exit(dbDisconnect(con))
  if (!"dc_meta" %in% dbListTables(con)) return(tibble())
  dbReadTable(con, "dc_meta") %>% arrange(desc(version)) %>% as_tibble()
}

# Push `new_df` in as the live table, snapshotting whatever was live before
# it into history first. Nothing is ever deleted.
save_new_version <- function(new_df, note = "Manual update") {
  con <- get_con(); on.exit(dbDisconnect(con))

  meta <- if ("dc_meta" %in% dbListTables(con)) {
    dbReadTable(con, "dc_meta")
  } else {
    tibble(version = integer(), timestamp = character(), note = character(), n_rows = integer())
  }

  if ("dc_current" %in% dbListTables(con)) {
    cur <- dbReadTable(con, "dc_current")
    next_v <- if (nrow(meta) == 0) 1L else max(meta$version) + 1L
    dbWriteTable(con, paste0("dc_history_v", next_v), cur, overwrite = TRUE)
    meta <- bind_rows(meta, tibble(
      version = next_v, timestamp = as.character(Sys.time()),
      note = paste0("Auto-snapshot before: ", note), n_rows = nrow(cur)
    ))
  }

  dbWriteTable(con, "dc_current", new_df, overwrite = TRUE)

  next_v2 <- max(meta$version, 0L) + 1L
  dbWriteTable(con, paste0("dc_history_v", next_v2), new_df, overwrite = TRUE)
  meta <- bind_rows(meta, tibble(
    version = next_v2, timestamp = as.character(Sys.time()), note = note, n_rows = nrow(new_df)
  ))
  dbWriteTable(con, "dc_meta", meta, overwrite = TRUE)
}

restore_version <- function(v) {
  con <- get_con()
  tbl_name <- paste0("dc_history_v", v)
  if (!tbl_name %in% dbListTables(con)) {
    dbDisconnect(con)
    return(FALSE)
  }
  old <- dbReadTable(con, tbl_name)
  dbDisconnect(con)
  save_new_version(old, note = paste("Restored from version", v))
  TRUE
}

us_state_codes <- c(
  "AL","AK","AZ","AR","CA","CO","CT","DE","FL","GA","HI","ID","IL","IN","IA",
  "KS","KY","LA","ME","MD","MA","MI","MN","MS","MO","MT","NE","NV","NH","NJ",
  "NM","NY","NC","ND","OH","OK","OR","PA","RI","SC","SD","TN","TX","UT","VT",
  "VA","WA","WV","WI","WY","DC","US","PA/GA"
)

# ---------- UI ----------

ui <- page_sidebar(
  title = tags$span(
    tags$img(src = "logo.png", height = "80px", style = "margin-right:10px; vertical-align:middle;",
              onerror = "this.style.display='none'")
  ),
  
  theme = bs_theme(
    version = 5,
    bg = "#0B0F14",
    fg = "#E2E8F0",
    primary = "#3B82F6",
    secondary = "#64748B",
    success = "#22C55E",
    warning = "#F59E0B",
    danger = "#EF4444",
    base_font = font_google("Inter"),
    heading_font = font_google("Inter")
  ) %>%
    bs_add_rules("
      html, body,
      .bslib-page-sidebar,
      .bslib-sidebar-layout,
      .bslib-sidebar-layout > .main,
      .tab-content,
      .container-fluid {
        background-color: #0B0F14 !important;
        color: #E2E8F0 !important;
      }

      .card, .card-header, .card-body {
        background-color: #111827 !important;
        color: #E2E8F0 !important;
        border-color: #1F2937 !important;
      }

      .bslib-value-box,
      .bslib-value-box .value-box-area,
      .bslib-value-box .value-box-showcase {
        background-color: #111827 !important;
        border-color: #1F2937 !important;
        color: #E2E8F0 !important;
      }

      .bslib-sidebar-layout > .sidebar {
        background-color: #111827 !important;
        color: #E2E8F0 !important;
      }

      .leaflet-container {
        background-color: #0B0F14 !important;
      }

      .dataTables_wrapper {
        background-color: #111827 !important;
        color: #E2E8F0 !important;
      }
      table.dataTable, table.dataTable td, table.dataTable th {
        background-color: #111827 !important;
        color: #E2E8F0 !important;
        border-color: #1F2937 !important;
      }
      table.dataTable tbody tr:hover {
        background-color: #1F2937 !important;
      }
      .dataTables_wrapper .dataTables_length,
      .dataTables_wrapper .dataTables_filter,
      .dataTables_wrapper .dataTables_info,
      .dataTables_wrapper .dataTables_paginate {
        color: #E2E8F0 !important;
      }
      .dataTables_wrapper .form-control {
        background-color: #0B0F14 !important;
        color: #E2E8F0 !important;
        border-color: #1F2937 !important;
      }
    "),

  sidebar = sidebar(
    width = 300,
    radioButtons("scope", "Scope", choices = c("US only" = "us", "Global" = "global"), selected = "us"),
    selectizeInput("state_filter", "State (US)", choices = NULL, multiple = TRUE,
                   options = list(placeholder = "All states")),
    conditionalPanel(
      condition = "input.scope == 'global'",
      selectizeInput("country_filter", "Country (non-US)", choices = NULL, multiple = TRUE,
                     options = list(placeholder = "All countries"))
    ),
    selectizeInput("city_filter", "City", choices = NULL, multiple = TRUE,
                   options = list(placeholder = "All cities")),
    selectizeInput("operator_filter", "Operator", choices = NULL, multiple = TRUE,
                   options = list(placeholder = "All operators")),
    sliderInput("capacity_filter", "Capacity (MW est.)", min = 0, max = 100,
                value = c(0, 100), step = 1),
    hr(),
    actionButton("reset_filters", "Reset filters", icon = icon("rotate-left"))
  ),

  navset_tab(
    nav_panel(
      "Map & Table",
      layout_columns(
        col_widths = c(4, 4, 4),
        value_box(title = "Operators in view", value = textOutput("vb_operators"), showcase = icon("building"), theme = "primary"),
        value_box(title = "Cities in view", value = textOutput("vb_cities"), showcase = icon("city"), theme = "secondary"),
        value_box(title = "Total capacity (MW est.)", value = textOutput("vb_capacity"), showcase = icon("bolt"), theme = "success")
      ),
      card(
        full_screen = TRUE,
        card_header("Map of Total Capacity (zoom in to split a cluster)"),
        leafletOutput("map", height = 520)
      ),
      card(
        card_header("Matching rows"),
        DTOutput("table")
      )
    ),
    nav_panel(
      "Version History",
      card(
        card_header("Load a new export"),
        p("Uploading a CSV here snapshots the current live data into history first, ",
          "then makes the new file live. Nothing is overwritten destructively."),
        fileInput("new_csv", "CSV file (same columns as the original export)", accept = ".csv"),
        textInput("version_note", "Note for this version", value = ""),
        actionButton("load_version", "Load as new version", icon = icon("upload"))
      ),
      card(
        card_header("History"),
        p("Restoring a version snapshots the current data first, then makes the chosen version live again \u2014 so a restore is itself reversible."),
        DTOutput("version_table")
      )
    )
  )
)

# ---------- Server ----------

server <- function(input, output, session) {

  raw_data <- reactiveVal(load_current())

  observeEvent(list(raw_data(), input$scope), {
    df <- raw_data()
    updateSelectizeInput(session, "state_filter", choices = sort(unique(df$State[df$is_us])), server = TRUE)
    updateSelectizeInput(session, "country_filter", choices = sort(unique(df$State[!df$is_us])), server = TRUE)
    df_scoped <- if (input$scope == "us") df %>% filter(is_us) else df
    max_cap <- max(df_scoped$Capacity_MW_est, na.rm = TRUE)
    updateSliderInput(session, "capacity_filter", max = ceiling(max_cap),
                       value = c(0, ceiling(max_cap)))
  }, ignoreNULL = FALSE)

  # Locations picked in either the State or Country box narrow the same
  # underlying State column, so they combine with OR, not AND.
  selected_locations <- reactive({
    if (input$scope == "us") return(input$state_filter)
    c(input$state_filter, input$country_filter)
  })

  observeEvent(list(raw_data(), input$scope, input$state_filter, input$country_filter), {
    df <- raw_data()
    if (input$scope == "us") df <- df %>% filter(is_us)
    locs <- selected_locations()
    if (length(locs) > 0) df <- df %>% filter(State %in% locs)
    freezeReactiveValue(input, "city_filter")
    updateSelectizeInput(session, "city_filter", choices = sort(unique(df$City_clean)), server = TRUE)
  }, ignoreNULL = FALSE)

  observeEvent(list(raw_data(), input$scope, input$state_filter, input$country_filter, input$city_filter), {
    df <- raw_data()
    if (input$scope == "us") df <- df %>% filter(is_us)
    locs <- selected_locations()
    if (length(locs) > 0) df <- df %>% filter(State %in% locs)
    if (length(input$city_filter) > 0) df <- df %>% filter(City_clean %in% input$city_filter)
    freezeReactiveValue(input, "operator_filter")
    updateSelectizeInput(session, "operator_filter", choices = sort(unique(df$Operator)), server = TRUE)
  }, ignoreNULL = FALSE)

  observeEvent(input$reset_filters, {
    updateSelectizeInput(session, "state_filter", selected = character(0))
    updateSelectizeInput(session, "country_filter", selected = character(0))
    updateSelectizeInput(session, "city_filter", selected = character(0))
    updateSelectizeInput(session, "operator_filter", selected = character(0))
    updateRadioButtons(session, "scope", selected = "us")
  })

  filtered <- reactive({
    df <- raw_data()
    if (input$scope == "us") df <- df %>% filter(is_us)
    locs <- selected_locations()
    if (length(locs) > 0) df <- df %>% filter(State %in% locs)
    if (length(input$city_filter) > 0) df <- df %>% filter(City_clean %in% input$city_filter)
    if (length(input$operator_filter) > 0) df <- df %>% filter(Operator %in% input$operator_filter)
    df <- df %>% filter(
      is.na(Capacity_MW_est) |
        (Capacity_MW_est >= input$capacity_filter[1] & Capacity_MW_est <= input$capacity_filter[2])
    )
    df
  })

  output$vb_operators <- renderText({ n_distinct(filtered()$Operator) })
  output$vb_cities     <- renderText({ n_distinct(filtered()$City_clean) })
  output$vb_capacity   <- renderText({
    total <- sum(filtered()$Capacity_MW_est, na.rm = TRUE)
    format(round(total), big.mark = ",")
  })

  output$map <- renderLeaflet({
    leaflet(options = leafletOptions(worldCopyJump = FALSE, minZoom = 2, maxZoom = 18)) %>%
      #addProviderTiles(providers$OpenStreetMap.Mapnik, options = providerTileOptions(noWrap = TRUE)) %>%
      addProviderTiles(providers$Esri.WorldGrayCanvas, options = providerTileOptions(noWrap = TRUE)) %>%
      setMaxBounds(lng1 = -180, lat1 = -85, lng2 = 180, lat2 = 85) %>%
      setView(lng = -98.5, lat = 39.5, zoom = 4)
  })

  observe({
    df <- filtered() %>% filter(!is.na(Latitude), !is.na(Longitude))
    leafletProxy("map", data = df) %>%
      clearMarkers() %>%
      clearMarkerClusters() %>%
      addCircleMarkers(
        lng = ~Longitude, lat = ~Latitude,
        radius = 8,
        fillOpacity = 0.8,
        color = "#60A5FA",
        fillColor = "#3B82F6",
        weight = 2,
        popup = ~paste0(
          "<b>", Operator, "</b><br>",
          City_clean, ", ", State, "<br>",
          "Capacity: ", ifelse(is.na(Capacity), "n/a", Capacity)
        ),
        clusterOptions = markerClusterOptions()
      )
  })

  output$table <- renderDT({
    filtered() %>%
      select(Operator, City = City_clean, State, Region, Capacity, `Capacity (MW est.)` = Capacity_MW_est) %>%
      arrange(Operator, State, City)
  }, options = list(pageLength = 15), rownames = FALSE)

  # ---- Version history tab ----

  refresh_trigger <- reactiveVal(0)

  output$version_table <- renderDT({
    refresh_trigger()
    versions <- list_versions()
    if (nrow(versions) == 0) return(datatable(tibble(Message = "No versions yet.")))
    versions %>%
      arrange(desc(version)) %>%
      select(Version = version, Timestamp = timestamp, Note = note, Rows = n_rows) %>%
      datatable(
        selection = "single", rownames = FALSE,
        options = list(pageLength = 10)
      )
  })

  observeEvent(input$load_version, {
    req(input$new_csv)
    new_df <- tryCatch(read_csv(input$new_csv$datapath, show_col_types = FALSE), error = function(e) NULL)
    if (is.null(new_df)) {
      showNotification("Could not read that CSV.", type = "error")
      return()
    }
    required_cols <- c("Operator", "City", "State")
    if (!all(required_cols %in% names(new_df))) {
      showNotification(
        paste("CSV must contain columns:", paste(required_cols, collapse = ", ")),
        type = "error"
      )
      return()
    }
    new_df <- new_df %>%
      mutate(
        Operator = str_trim(Operator),
        City = str_trim(City),
        State = str_trim(State),
        Region = if ("Region" %in% names(new_df)) str_trim(Region) else NA_character_,
        Capacity = if ("Capacity" %in% names(new_df)) Capacity else NA_character_,
        Capacity_MW_est = if ("Capacity_MW_est" %in% names(new_df)) as.numeric(Capacity_MW_est) else NA_real_,
        City_clean = if ("City_clean" %in% names(new_df)) str_trim(City_clean) else City,
        Latitude = if ("Latitude" %in% names(new_df)) as.numeric(Latitude) else NA_real_,
        Longitude = if ("Longitude" %in% names(new_df)) as.numeric(Longitude) else NA_real_,
        is_us = State %in% us_state_codes
      )
    note <- if (nchar(input$version_note) > 0) input$version_note else "Manual update via app"
    save_new_version(new_df, note = note)
    raw_data(load_current())
    refresh_trigger(refresh_trigger() + 1)
    showNotification("New version loaded and now live.", type = "message")
  })

  observeEvent(input$version_table_rows_selected, {
    sel <- input$version_table_rows_selected
    req(sel)
    versions <- list_versions() %>% arrange(desc(version))
    v <- versions$version[sel]
    showModal(modalDialog(
      title = paste("Restore version", v, "?"),
      paste0("This will make version ", v, " live again. The current data will be ",
             "snapshotted first, so this is safe to undo."),
      footer = tagList(
        modalButton("Cancel"),
        actionButton("confirm_restore", "Restore", class = "btn-danger")
      ),
      easyClose = TRUE
    ))
    session$userData$pending_restore <- v
  })

  observeEvent(input$confirm_restore, {
    v <- session$userData$pending_restore
    req(v)
    restore_version(v)
    raw_data(load_current())
    refresh_trigger(refresh_trigger() + 1)
    removeModal()
    showNotification(paste("Restored version", v, "and set it live."), type = "message")
  })
}

shinyApp(ui, server)
