library(shiny)
library(shinyjs)
library(shinythemes)
library(sf)
library(readxl)
library(dplyr)
library(tidyr)
library(stringr)
library(leaflet)
library(viridisLite)
library(htmlwidgets)
library(ggplot2)

sf::sf_use_s2(FALSE)

# --------------------------- DATA PREP ---------------------------------------

# --- ADM2 SHAPEFILE (GROUND TRUTH FOR PCODES + GEOMETRY) ---------------------
adm2 <- st_read("input/Admin2/SDN_Admin2.shp", quiet = TRUE) %>%
  mutate(
    ADM2_EN = case_when(
      ADM2_EN == "Ar Rahad" & ADM1_EN == "Gedaref" ~ "Ar Rahad - Gf",
      ADM2_EN == "Ar Rahad" & ADM1_EN == "North Kordofan" ~ "Ar Rahad - NK",
      TRUE ~ ADM2_EN  # Keep the original value for other rows
    )
  )%>%
  st_make_valid() %>%
  st_transform(4326)

ana_summary<-read_xlsx("input/99_ANA_Flagging_Deep_Analysis.xlsx", sheet = "STEP 5 Flagging Chart")  %>%
  select(-c("Conclusion Long","Conclusion Short"))
  

# --------------------------- LONG ANA DATA (SHEET 2) -------------------------
ana_long <- read_xlsx("input/99_ANA_Flagging_Deep_Analysis.xlsx", sheet = "Final Deep Dive Results_long") %>%
  mutate(
    indicator = str_squish(indicator),
    value     = str_squish(value),
    indicator = paste0(toupper(substr(indicator, 1, 1)), substr(indicator, 2, nchar(indicator))),  # Capitalize only the first letter
    value     = paste0(toupper(substr(value, 1, 1)), substr(value, 2, nchar(value)))  # Capitalize only the first letter
  )

ana_base_data <- read_xlsx("input/99_ANA_Flagging_Deep_Analysis.xlsx", sheet = "Source Data Flagged") %>%
  mutate(`Indicator Value`=round(`Indicator Value`,2)) %>%
  rename( `Administrative Level P-Code`=`Unit of Analysis`,
          `Administrative Level Label`=`Unit of Analysis Label`) %>%
  arrange(`Administrative Level Label`,System )

ana_long_download<- ana_long  %>%
  mutate(indicator=ifelse(indicator=="Conclusion" & System=="Deep Dive Summary", "to_remove", indicator),
         indicator=ifelse(indicator=="Conclusion summary" & System=="Deep Dive Summary", "to_remove", indicator)) %>%
  filter(!(indicator %in% c("Deep dive decision HQ standardization", "SUMMARY PRELIMINARY FLAG CONCLUSION", "to_remove"))) 



# --------------------------- MAP ATTRIBUTES FROM LONG ------------------------
# Build deep_dive + main_driver per ADM2 from long format
map_attrs <- ana_long %>%
  filter(indicator %in% c("Deep dive decision", "Main driver",	
                          "Deep dive decision HQ standardization")) %>%
  mutate(
    indicator_clean = case_when(
      indicator == "Deep dive decision" ~ "deep_dive",
      indicator == "Main driver"        ~ "main_driver",
      indicator == "Deep dive decision HQ standardization"  ~ "deep_dive_flat",
      TRUE                              ~ NA_character_
    )
  ) %>%
  filter(!is.na(indicator_clean)) %>%
  select(
    ADM2_PCODE = `Administrative Level P-Code`,
    ADM2_EN    = `Administrative Level Label`,
    State,
    indicator_clean,
    value
  ) %>%
  distinct(ADM2_PCODE, indicator_clean, .keep_all = TRUE) %>%
  pivot_wider(
    names_from  = indicator_clean,
    values_from = value
  )

# Join onto adm2 geometries
summary_data <- adm2 %>%
  left_join(map_attrs) %>%
  arrange(ADM1_EN, ADM2_EN)

# --------------------------- PALETTE: DEEP DIVE NUANCED -----------------------

main_roem_levels <- sort(unique(na.omit(summary_data$deep_dive)))
n_drivers_roem <- length(main_roem_levels)


roem_deep_dive_palette <- if (n_drivers_roem <= 9) {
  c(
    "#FFD700",
    "#FFA500",
    "#8B0000",
    "#FF4500",
    "#B22222",
    "#f0f7f0"
  )[seq_len(n_drivers_roem)]
} else {
  turbo(n_drivers_roem )
}

driver_colors_roem <- setNames(roem_deep_dive_palette, main_roem_levels )

pal_deepdive <- colorFactor(
  palette  = driver_colors_roem,
  domain   = main_roem_levels,
  na.color = "#f0f0f0"
)

# --------------------------- PALETTE: DEEP DIVE NUANCED -----------------------

main_roem_levels_levels <- c("Acute Needs", "Acute Needs (Suspect RoEM)", "RoEM", "RoEM (Suspect EM)", "EM")


roem_deep_dive_palette_levels <- if (n_drivers_roem <= 9) {
  c(
    "#FFD700",
    "#FFA500",
    "#8B0000",
    "#FF4500",
    "#B22222",
    "#f0f7f0"
  )[seq_len(n_drivers_roem)]
} else {
  turbo(n_drivers_roem )
}

driver_colors_roem_levels <- setNames(roem_deep_dive_palette_levels, main_roem_levels_levels )

pal_deepdive_levels <- colorFactor(
  palette  = driver_colors_roem_levels,
  domain   = main_roem_levels_levels,
  na.color = "#f0f0f0"
)

# --------------------------- PALETTE: DEEP DIVE FLAT --------------------------

main_roem_flat_levels <- sort(unique(na.omit(summary_data$deep_dive_flat)))
n_drivers_roem_flat <- length(main_roem_flat_levels)


roem_palette_flat <- if (n_drivers_roem_flat <= 9) {
  c(
    "#FFD700",
    "#8B0000",
    "#FF4500"
  )[seq_len(n_drivers_roem_flat)]
} else {
  turbo(n_drivers_roem_flat )
}

driver_colors_roem_flat <- setNames(roem_palette_flat, main_roem_flat_levels )

pal_deepdive_flat <- colorFactor(
  palette  = driver_colors_roem_flat,
  domain   = main_roem_flat_levels,
  na.color = "#f0f0f0"
)


# --------------------------- PALETTE: MAIN DRIVER ----------------------------
main_driver_levels <- sort(unique(na.omit(summary_data$main_driver)))
n_drivers <- length(main_driver_levels)

driver_palette <- if (n_drivers <= 9) {
  c(
    "#99000d", "#1A9850", "#FEE08B", "#D73027", "#542788", "#4575B4", "#313695", "#542788"
  )[seq_len(n_drivers)]
} else {
  turbo(n_drivers)
}

driver_colors <- setNames(driver_palette, main_driver_levels)

pal_driver <- colorFactor(
  palette  = driver_colors,
  domain   = main_driver_levels,
  na.color = "#f0f0f0"
)





# --------------------------- DONUT DATA WITHIN (FROM LONG) --------------------------
conclusion_colors <- c(
  "C1: Systemic Failure"          = "#FF9F9A",
  "C2: Pockets of severity"       = "#FFE4A3",
  "C3: Mitigation/no acute needs" = "#C9F2A3",
  "Inconclusive"                  = "#D3D3D3"
)

deep_conclusion <- ana_long %>%
  filter(indicator == "Conclusion") %>%
  mutate(
    conclusion_clean = case_when(
      value == "C1: Systemic Failure"             ~ "C1: Systemic Failure",
      value == "C2: Pockets of severity"          ~ "C2: Pockets of severity",
      value %in% c("C3: Mitigation/no acute needs",
                   "C3: Migation/no acute needs") ~ "C3: Mitigation/no acute needs",
      value == "Inconclusive"          ~ "Inconclusive",
      TRUE ~ NA_character_
    )
  ) %>%
  filter(!is.na(conclusion_clean)) %>%
  select(`Administrative Level Label`, conclusion = conclusion_clean)

deep_shares <- deep_conclusion %>%
  count(`Administrative Level Label`, conclusion, name = "n") %>%
  group_by(`Administrative Level Label`) %>%
  complete(conclusion = names(conclusion_colors), fill = list(n = 0)) %>%
  mutate(share = n / sum(n)) %>%
  ungroup()



# --------------------------- DONUT DATA ACROSS (FROM LONG) --------------------------
conclusion_colors2 <- c(
  "C1: Strong Interaction"          = "#FF9F9A",
  "C2: Limited or indirect Interaction"       = "#FFE4A3",
  "C3: No interaction" = "#C9F2A3"
)

deep_conclusion2 <- ana_long %>%
  filter(indicator == "Conclusion") %>%
  mutate(
    conclusion_clean = case_when(
      value == "C1: Strong Interaction"             ~ "C1: Strong Interaction",
      value == "C2: Limited or indirect Interaction"~ "C2: Limited or indirect Interaction",
      value %in% c("C3: No interaction")            ~ "C3: No interaction",
      TRUE ~ NA_character_
    )
  ) %>%
  filter(!is.na(conclusion_clean)) %>%
  select(`Administrative Level Label`, conclusion = conclusion_clean)

deep_shares2 <- deep_conclusion2 %>%
  count(`Administrative Level Label`, conclusion, name = "n") %>%
  group_by(`Administrative Level Label`) %>%
  complete(conclusion = names(conclusion_colors), fill = list(n = 0)) %>%
  mutate(share = n / sum(n)) %>%
  ungroup()

# --------------------------- SELECTIZE CHOICES -------------------------------
adm2_nog <- summary_data %>%
  st_set_geometry(NULL) %>%
  arrange(ADM1_EN, ADM2_EN)

adm2_choices <- split(adm2_nog$ADM2_EN, adm2_nog$ADM1_EN)

# --------------------------- CENTROIDS ---------------------------------------

cent_sf <- st_point_on_surface(summary_data)
cent_coords <- st_coordinates(cent_sf)

summary_centroids <- data.frame(
  ADM2_EN = summary_data$ADM2_EN,
  lon     = cent_coords[, "X"],
  lat     = cent_coords[, "Y"],
  stringsAsFactors = FALSE
)


# --------------------------- NULL COALESCE -----------------------------------
`%||%` <- function(a, b) if (!is.null(a)) a else b

# ================================ UI ==========================================

ui <- shiny::bootstrapPage(
  shinyjs::useShinyjs(),
  tags$head(
    tags$link(rel = "stylesheet", href = "https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0-beta3/css/all.min.css"),
    tags$link(rel = "stylesheet", href = "https://cdnjs.cloudflare.com/ajax/libs/font-awesome/5.15.3/css/all.min.css"),
    tags$script(HTML("
  Shiny.addCustomMessageHandler('focusLocSelect', function(msg) {
    var el = $('#loc_select')[0];
    if (el && el.selectize) {
      el.selectize.focus();
    }
  });

  window._anaMap = window._anaMap || null;
  window._anaConnectorCfg = window._anaConnectorCfg || null;

  function anaRedrawConnector() {
    var map = window._anaMap;
    var cfg = window._anaConnectorCfg;
    if (!map) return;

    if (map._connectorLayer) {
      map.removeLayer(map._connectorLayer);
      map._connectorLayer = null;
    }
    if (map._connectorDot) {
      map.removeLayer(map._connectorDot);
      map._connectorDot = null;
    }

    if (!cfg || !cfg.active || cfg.lat === undefined || cfg.lng === undefined) return;

    var panel = document.getElementById('info-panel');
    if (!panel) return;

    var mapContainer = map.getContainer();
    var mapRect = mapContainer.getBoundingClientRect();
    var rect = panel.getBoundingClientRect();
    if (rect.width === 0 || rect.height === 0) return;

    var p0 = map.latLngToContainerPoint([cfg.lat, cfg.lng]);

    var anchorX = rect.left - mapRect.left;
    var anchorY = rect.top - mapRect.top + rect.height / 2;
    var p3 = L.point(anchorX, anchorY);

    var dx = p3.x - p0.x;
    var dy = p3.y - p0.y;

    if (Math.abs(dx) < 10) {
      if (!map.getPane('connectorPane')) {
        map.createPane('connectorPane');
        var paneStraight = map.getPane('connectorPane');
        paneStraight.style.zIndex = 650;
        paneStraight.style.pointerEvents = 'none';
      }

      var straightCoords = [
        map.containerPointToLatLng(p0),
        map.containerPointToLatLng(p3)
      ];

      var straight = L.polyline(straightCoords, {
        pane: 'connectorPane',
        color: '#666666',
        weight: 0.8,
        opacity: 0.45,
        lineJoin: 'round',
        lineCap: 'round',
        smoothFactor: 2
      }).addTo(map);

      var dotStraight = L.circleMarker([cfg.lat, cfg.lng], {
        pane: 'connectorPane',
        radius: 2.5,
        color: '#666666',
        weight: 0.8,
        opacity: 0.7,
        fillColor: '#ffffff',
        fillOpacity: 0.9
      }).addTo(map);

      map._connectorLayer = straight;
      map._connectorDot   = dotStraight;
      return;
    }

    var c1 = L.point(
      p0.x + dx * 0.33,
      p0.y
    );
    var c2 = L.point(
      p3.x - dx * 0.33,
      p3.y
    );

    var coords = [];
    var steps = 120;
    for (var i = 0; i <= steps; i++) {
      var t = i / steps;
      var oneMinusT = 1 - t;

      var x =
        oneMinusT * oneMinusT * oneMinusT * p0.x +
        3 * oneMinusT * oneMinusT * t * c1.x +
        3 * oneMinusT * t * t * c2.x +
        t * t * t * p3.x;

      var y =
        oneMinusT * oneMinusT * oneMinusT * p0.y +
        3 * oneMinusT * oneMinusT * t * c1.y +
        3 * oneMinusT * t * t * c2.y +
        t * t * t * p3.y;

      var latlng = map.containerPointToLatLng([x, y]);
      coords.push(latlng);
    }

    if (!map.getPane('connectorPane')) {
      map.createPane('connectorPane');
      var pane = map.getPane('connectorPane');
      pane.style.zIndex = 650;
      pane.style.pointerEvents = 'none';
    }

    var line = L.polyline(coords, {
      pane: 'connectorPane',
      color: '#666666',
      weight: 0.8,
      opacity: 0.45,
      lineJoin: 'round',
      lineCap: 'round',
      smoothFactor: 2
    }).addTo(map);

    var dot = L.circleMarker([cfg.lat, cfg.lng], {
      pane: 'connectorPane',
      radius: 2.5,
      color: '#666666',
      weight: 0.8,
      opacity: 0.7,
      fillColor: '#ffffff',
      fillOpacity: 0.9
    }).addTo(map);

    map._connectorLayer = line;
    map._connectorDot   = dot;
  }

  Shiny.addCustomMessageHandler('updateConnector', function(msg) {
    window._anaConnectorCfg = msg;

    function ensureDraw(attempts) {
      attempts = attempts || 0;
      if (!window._anaMap) {
        if (attempts < 20) {
          setTimeout(function() { ensureDraw(attempts + 1); }, 100);
        }
        return;
      }
      anaRedrawConnector();
      if (!window._anaMap._connectorMoveHandlerAttached) {
        window._anaMap.on('zoomend moveend', function() {
          anaRedrawConnector();
        });
        window._anaMap._connectorMoveHandlerAttached = true;
      }
    }

    ensureDraw(0);
  });
")),########################## STYLING #################################
    tags$style(HTML("
/* ================== RIGHT INFO PANEL (SIDEPANEL) ================== */
#info-panel {
  position: absolute;
  top: 76px;
  right: 22px;
  width: 340px;
  z-index: 1000;
  margin-top: 10px;
  height: calc(100vh - 160px);  /* Dynamic height based on screen size */
  display: flex;
  flex-direction: column;
  padding-right: 8px;
}

/* Custom scrollbar styling */
#info-panel::-webkit-scrollbar {
  width: 4px; /* Thinner scrollbar for WebKit browsers (Chrome/Safari) */
}

#info-panel::-webkit-scrollbar-thumb {
  background-color: #ccc;
  border-radius: 10px;
}

#info-panel::-webkit-scrollbar-track {
  background-color: #e0e0e0;
  border-radius: 10px;
}

/* Panel content styling */
#info-panel .panel {
  padding: 10px 12px 8px 12px;
  font-size: 12px;
  line-height: 1.3;
  flex-grow: 1;  /* Ensures panel grows with content */
  display: flex;
  flex-direction: column;
  gap: 8px;  /* Space between elements */
}

/* Adjusting the scrollable area */
#info-panel .panel > div {
  overflow-y: auto;
  padding-right: 14px;  /* Space between content and scrollbar */
  flex-grow: 1;
  max-height: 100%;
  overflow: hidden;  /* Prevent scrolling for the title */
}

/* Styling for the title (header) */
#info-panel h4 {
  margin-top: 0;
  margin-bottom: 0px;
  font-weight: 700;
  font-size: 14px;
  overflow: hidden;  /* Prevent title from scrolling */
}

/* ================== SLIDER THUMB ================== */
input[type='range']::-webkit-slider-thumb {
  -webkit-appearance: none;
  height: 20px;
  width: 20px;
  border-radius: 50%;
  background: #333;
  margin-top: -6px;  /* Adjust thumb vertical position */
}

/* Sidebar table adjustments */
#sidebar_table {
  margin-top: 10px;
  margin-bottom: 20px;
}

/* Side-label styling */
.side-label {
  margin-top: 10px;
  font-weight: bold;
  font-size: 14px;
}

/* Horizontal rule styling inside panels */
.panel hr {
  margin: 6px 0;
  border-color: #ccc;
  border-width: 1px;
}

/* ================== GLOBAL RESET ================== */
html, body {
  margin: 0;
  padding: 0;
  scroll-behavior: smooth;
}

body {
  padding-top: 0 !important;
}

/* ================== FULL-PAGE MAP SHELL ================== */
#clusterWrap {
  position: fixed;
  top: 50px;
  left: 0;
  right: 0;
  bottom: 0;
}

#clusterWrap .leaflet-container {
  height: 100% !important;
  width: 100% !important;
}

/* ================== LEFT CONTROL PANEL (MAP) ================== */
#cluster-controls {
  position: absolute;
  top: 76px;
  left: 22px;
  width: 260px;
  z-index: 1000;
}

#cluster-controls .panel {
  padding: 10px;
}

#cluster-controls h5 {
  margin: 8px 0 6px;
  font-weight: 700;
}

#cluster-controls .help {
  font-size: 11px;
  color: #666;
  margin-bottom: 6px;
}

#cluster-controls .form-group {
  margin-bottom: 8px;
}

#cluster-controls .selectize-control.single .selectize-input {
  min-height: 30px;
  padding: 4px 8px;
  font-size: 12px;
}

#cluster-controls .selectize-dropdown {
  font-size: 12px;
}

/* ================== SIDEBAR TEXT STYLING ================== */
.side-title {
  font-size: 16px;
  font-weight: 700;
  margin-bottom: 6px;
}

.fade-state {
  opacity: 0.6;
  font-weight: 600;
}

.side-label {
  font-weight: 600;
  margin-top: 6px;
}

.side-value {
  margin-left: 4px;
  font-size: 12px;
}

/* ================== TOGGLE SWITCH ================== */
.switch {
  position: relative;
  display: inline-block;
  width: 26px;
  height: 14px;
  margin-top: 7px;      /* move DOWN */
}

.switch input {
  opacity: 0;
  width: 0;
  height: 0;
}

.slider {
  position: absolute;
  cursor: pointer;
  top: 0;
  left: 0;
  right: 0;
  bottom: 0;
  background-color: #ccc;
  transition: .25s;
  border-radius: 14px;    /* adjust to smaller height */
}

.slider:before {
  position: absolute;
  content: '';
  height: 10px;   /* smaller circle */
  width: 10px;
  left: 2px;
  bottom: 2px;
  background-color: white;
  transition: .25s;
  border-radius: 50%;
}

/* ON = black */
input:checked + .slider {
  background-color: #000000 !important;
}

/* Circle slide */
input:checked + .slider:before {
  transform: translateX(12px);  /* adjusted for smaller switch */
}

/* ================== HOTSPOTS LAYOUT (NAV + CONTENT) ================== */

/* Layout wrapper: sidebar + main content */
.hotspots-layout {
  display: flex;
  align-items: flex-start;
}

/* Left side panel */
.hotspots-sidebar {
  flex: 0 0 260px;
  margin-right: 24px;
  position: sticky;
  top: 100px;
  align-self: flex-start;
}

/* Main content adapts to remaining space */
.hotspots-main {
  flex: 1;
}

/* Contents title */
.hotspots-nav-title {
  font-weight: 600;
  font-size: 14px;
  text-transform: uppercase;
  letter-spacing: 0.08em;
  color: #000000;
  margin-bottom: 10px;
}

/* Contents list */
.hotspots-nav-links {
  list-style: none;
  margin: 0;
  padding: 0;
}

.hotspots-nav-links li {
  margin-bottom: 4px;
}

.hotspots-nav-links li a {
  display: block;
  text-decoration: none;
  color: #333333;
  background-color: #ffffff;
  border-radius: 4px;
  padding: 6px 8px;
  border: 1px solid transparent;
  font-size: 0.9rem;
  line-height: 1.1;
  font-weight: 600;
  text-transform: none;
}

.hotspots-nav-links li a:hover {
  background-color: #f4f7fb;
  border-color: #d0d7e2;
}

.hotspots-nav-links li a.active-link {
  background-color: #155e8f;
  border-color: #155e8f;
  color: #ffffff;
}

/* Keep nice offset for anchor jumps */
.hotspot-section {
  scroll-margin-top: 90px;
}

/* Make layout adaptive when the window is narrow */
.hotspots-layout {
  display: flex;
  align-items: flex-start;
}

.hotspots-sidebar {
  flex: 0 0 260px;
  margin-right: 24px;
  position: sticky;      /* <- this makes it move with scroll */
  top: 100px;            /* distance from top (below navbar) */
  align-self: flex-start;
}

.hotspots-sidebar-box {
  background-color: #f5f5f5;      /* light grey */
  border-radius: 10px;            /* rounded edges */
  padding: 14px 16px;
}

.hotspots-main {
  flex: 1;
  /* optional: small side padding so it doesn’t feel cramped */
  padding-right: 20px;
}
}
/* ================== DOWNLOAD BUTTON (LOCALITY DATA) ================== */
#download_locality_data {
  background-color: #f2f2f2;
  color: #4a4a4a;
  border-radius: 999px;
  border: none;
  padding: 8px 18px;
  font-size: 12px;
  margin-top: 4px;
}

#download_locality_data:hover {
  background-color: #e0e0e0;
}

/* ================== DT ROW HEIGHT (ANA PREVIEW) ================== */
#ana_preview table.dataTable tbody td {
  padding-top: 6px !important;
  padding-bottom: 6px !important;
  line-height: 1.1 !important;
}

#ana_preview table.dataTable tbody tr {
  height: 32px !important;
}
"))),
  navbarPage(
    id="navbarPage",
    theme = shinythemes::shinytheme("flatly"),
    collapsible = TRUE,
    title = div(HTML('<a style="text-decoration:none;cursor:default;color:#FFFFFF;" class="active" href="#">REACH - Sudan Needs Assessment</a>')),
    windowTitle = "REACH - Sudan Needs Assessment",
    position = "fixed-top",
    
    ################################# RESULT MAPPING #############################
    
    tabPanel(
      "Result Mapping",
      div(
        id = "clusterWrap",
        tags$head(includeCSS("styles.css")),
        leafletOutput("map", width = "100%", height = "100%"),
        
        # Left controls (Updated to match right panel)
        absolutePanel(
          id = "cluster-controls", class = "panel panel-default",
          fixed = TRUE, draggable = TRUE,
          width = 265,
          top = 76, left = 10,
          
          div(
            class = "panel",
            style = "
      display: flex; 
      flex-direction: column; 
      gap: 6px;                /* tighter spacing */
      padding: 6px 8px;        /* small padding */
    ",
            
            h5(
              tags$i(class = "fa fa-layer-group"), " Map",
              style = "margin: 0; padding: 0; font-size: 13px;"
            ),
            
            div(
              class = "help",
              "Select what data you want to visualize:",
              style = "font-size:11px; margin-top:0; margin-bottom:2px;"
            ),
            
            selectInput(
              "map_type", label = NULL,
              choices = c(
                "Acute Needs Analysis (Detailed)" = "deepdive",
                "Acute Needs Analysis (Simplified)" = "flat",
                "Main driver of Risk of Excess Mortality" = "driver"
              ),
              selected = "deepdive",
              width = "100%"
            ),
            
            h5(
              tags$i(class = "fa fa-map-marker-alt"), " Locality",
              style = "margin: 4px 0 0 0; padding: 0; font-size: 13px;"
            ),
            
            div(
              class = "help",
              "Type/Select the locality you want to display additional information for:",
              style = "font-size:11px; margin-top:0; margin-bottom:2px;"
            ),
            
            selectizeInput(
              "loc_select",
              label = NULL,
              choices  = c("", adm2_choices),
              selected = "",
              multiple = FALSE,
              options = list(
                placeholder = "Please type/select Locality",
                openOnFocus = FALSE,
                create = FALSE,
                maxItems = 1
              )
            ),
            
            # Toggle row
            div(
              style = "display:flex; align-items:center; 
               justify-content:space-between; 
               margin-top:4px; margin-bottom:0;",
              tags$label("Show Map Description",
                         style="font-size:11px; font-weight:bold; margin:0;"
              ),
              
              div(   # NEW WRAPPER
                style = "margin-right: 90px;",   # shifts LEFT
                tags$label(
                  class = "switch",
                  tags$input(id="toggle_map_desc", type="checkbox"),
                  tags$span(class="slider round")
                )
            )
            ),
            
            shinyjs::hidden(
              div(
                id = "map_description_box",
                uiOutput("map_type_text"),
                style = "
          padding: 6px;
          border-left: 2px solid #ccc;
          margin-top:4px;
          max-height: 160px;
          overflow-y: auto;
          padding-right: 6px;
          font-size:11px;
          line-height:1.3;
        "
              )
            )
          )
        ),
        
        # Right info panel (sidepanel) - Same characteristics as the left panel
        shinyjs::hidden(
          absolutePanel(
            id = "info-panel", class = "panel panel-default",
            fixed = TRUE, draggable = FALSE,
            width = 340,
            top = 76, right = 22,
            div(
              class = "panel",
              style = "display: flex; flex-direction: column; gap: 8px; height: calc(100vh - 160px);",  # Same height and flex layout as left panel
              
              # Fixed height for header
              tags$div(style = "margin-top: 5px;"),  # Space above the header
              uiOutput("sidebar_header", style = "flex-shrink: 0; height: 25px; overflow: hidden;"),  # Same header size and behavior
              
              # Horizontal line between header and content
              tags$hr(style = "margin:6px 0 6px 0;"),
              
              # Scrollable content section (with defined height and overflow)
              div(
                style = "flex-grow: 1; overflow-y: auto; padding-bottom: 10px; height: calc(100vh - 160px);",  # Scrollable content, same height
                uiOutput("sidebar_table"),
                tags$hr(style = "margin:8px 0 4px 0;"),
                div(class = "side-label", "Within-System Results"),
                
                # Mini donut plots with fixed height and centered alignment
                div(
                  style = "flex-shrink: 0; display: flex; justify-content: center; align-items: center; margin-bottom: 10px;",
                  plotOutput("mini_donut", height = "80px", width = "230px")
                ),
                
                uiOutput("within_system_conclusions"),
                
                tags$hr(style = "margin:8px 0 4px 0;"),
                div(class = "side-label", "Across-System Results:"),
                
                div(
                  style = "flex-shrink: 0; display: flex; justify-content: center; align-items: center; margin-bottom: 10px;",
                  plotOutput("mini_donut2", height = "80px", width = "230px")
                ),
                
                uiOutput("across_system_conclusions")
              )
            )
            ))

    )
  ),
  
  ################################# LOCALITY REPORTS ###########################
  
  
  tabPanel(
    "Locality Reports",
    div(
      class = "content-wrapper",
      style = "margin-top: 70px;",  # Space to avoid navbar overlap
      tags$head(includeCSS("styles.css")),
      
      # Left panel: Selectize input for locality
      div(
        class = "row",
        style = "display: flex; justify-content: center; align-items: flex-start;",
        div(
          class = "col-md-4",
          style = "padding-right: 20px;",
          
          div(
            class = "help",
            style = "font-size: 12px;",  # Set font size to 12px or any other size you prefer
            "Select the Locality you want to output the report for:"
          ),
          
          # Selectize Input for locality selection
          selectizeInput(
            "locality_report", 
            label = "", 
            choices = c("", adm2_choices),  # Add an empty string as the first choice
            selected = "Khartoum",  # No default selection
            multiple = FALSE,
            options = list(
              placeholder = "Please type/select Locality",
              openOnFocus = FALSE,
              create = FALSE,
              maxItems = 1
            )
          ),
          
          # Small map for the selected locality (below the selectize input)
          leafletOutput("small_map", height = "300px"),
          tags$br(),
          downloadButton(
            outputId = "download_locality_data",
            label = "Download Locality Data",
            style = "
    border-radius: 25px;
    padding: 10px 20px;"
          )
          
          
        ),
        
        # Right panel: Display locality info (header and conclusions)
        div(
          class = "col-md-8",
          
          # Locality and State Header
          uiOutput("locality_report_header"),
          
          uiOutput("locality_report_details"),
          
          div(
            class = "panel-group", 
            id = "accordion",
            
            # First collapsible div for system report details
            div(
              class = "panel panel-default", 
              style = "border-radius: 10px; margin-bottom: 20px;",  # Rounded corners and space between sections
              `data-toggle` = "collapse", 
              `data-target` = "#collapseOne",  # Make the whole panel clickable
              div(
                class = "panel-heading", 
                h6(
                  class = "panel-title", 
                  style = "font-weight: bold; display: inline-block; width: 90%;", 
                  "Core Systems Report"
                ),
                a(
                  href = "#collapseOne", 
                  class = "accordion-toggle", 
                  `data-toggle` = "collapse", 
                  style = "float: right; margin-top: 5px; color: #000000;",  # Adjusted position and color
                  icon("chevron-down")  # You can replace this with an appropriate icon
                )
              ),
              div(
                id = "collapseOne", 
                class = "panel-collapse collapse", 
                div(class = "panel-body", uiOutput("system_report_details"))
              )
            ),
            
            # Second collapsible div for across system report details
            div(
              class = "panel panel-default", 
              style = "border-radius: 10px;",  # Rounded corners
              `data-toggle` = "collapse", 
              `data-target` = "#collapseTwo",  # Make the whole panel clickable
              div(
                class = "panel-heading", 
                h6(
                  class = "panel-title", 
                  style = "font-weight: bold; display: inline-block; width: 90%;", 
                  "Cross-System Interactions Report"
                ),
                a(
                  href = "#collapseTwo", 
                  class = "accordion-toggle", 
                  `data-toggle` = "collapse", 
                  style = "float: right; margin-top: 5px; color: #000000;",  # Adjusted position and color
                  icon("chevron-down")  # You can replace this with an appropriate icon
                )
              ),
              div(
                id = "collapseTwo", 
                class = "panel-collapse collapse", 
                div(class = "panel-body", uiOutput("across_system_report_details"))
              )
            )
          )
          
          
        )
      )
    )
  ),

  
  ################################# HOTSPOT REVIEW #############################
  
  tabPanel(
    "Hotspots Review",
    div(
      class = "content-wrapper",
      style = "margin-top: 70px;",
      tags$head(includeCSS("styles.css")),
      
      # LAYOUT: sidebar on the left, text on the right
div(
  class = "hotspots-layout",
  
  # LEFT SIDE PANEL
  div( style = "margin-top: 180px;",
    class = "hotspots-sidebar",          # <- column
    div(
      class = "hotspots-sidebar-box",    # <- black box inside column
      div(class = "hotspots-nav-title", "Contents"),
      tags$ul(
        class = "hotspots-nav-links",
        tags$li(a(href = "#dilling_kadugli",  "Dilling & Kadugli (Kordofan)")),
        tags$li(a(href = "#al_fasher",        "Al Fasher (North Darfur)")),
        tags$li(a(href = "#umdaj_bara",       "Um Daj Ahmed & Bara (North Kordofan)")),
        tags$li(a(href = "#khartoum_health",  "Khartoum & health collapse")),
        tags$li(a(href = "#western_bluenile", "Western Blue Nile")),
        tags$li(a(href = "#buram_assunta",    "Buram & As Sunta (South Darfur)")),
        tags$li(a(href = "#um_dukhun",        "Um Dukhun (Central Darfur)")),
        tags$li(a(href = "#habila",           "Habila (West Darfur)")),
        tags$li(a(href = "#reifi_telkok",     "Reifi Telkok (Kassala)"))
      )
    )
  ),
        

             # Main content shifted right so it doesn't sit under the nav box
        div(
               class = "hotspots-main",
               style = "margin-top: 0px; margin-left: 0px;",  # <- push content to the right
               # ----------------- DILLING & KADUGLI -----------------
               div(
                 id = "dilling_kadugli",
                 class = "hotspot-section",
                 
                 tags$h2(tags$b("Hotspots Needs Review")),
                 
                 fluidRow(
                   column(
                     6,
                     div(
                       style = "display: flex; align-items: flex-start; height: 100%; padding-top: 130px; padding-left: 0px",
                       img(src = "kadugli.png", style = "width: 100%; height: 100%")
                     )
                   ),
                   column(
                     6,
                     style = "padding-top:20px;",
                     div(class = "title", tags$h3(tags$b("Dilling and Kadugli (Kordofan Region)"))),
                     br(), br(),
                     p("Protracted siege and instability have driven catastrophic food consumption gaps and severe malnutrition in both localities. Both the IPC and FSLA classify these localities as at Risk of Famine, with plausible evidence of famine in Kadugli."),
                     p("Damaged water infrastructure, long queues, and physical barriers to water access indicate poor access to water, and the population might resort to using unimproved water sources as a result. Inadequate water consumption together with unhygienic sanitation will most likely raise the chance of contracting water-borne diseases like cholera, which has a high incidence in the region and extremely high fatality rates according to PAMI."),
                     p("This widespread risk of contracting diseases is exacerbated by non-functioning PHCs, which are in any case not accessible to most of the population."),
                     p("All primary systems seem to strongly interact to create a humanitarian crisis of catastrophic proportions. Destitution spreads to neighboring localities in Kordofan, which are mostly classified as at risk of famine by both IPC and FSLA. Together with the Al Fasher cluster of localities, these seem to be the most critical crisis hotspots.")
                   )
                 ),
                 tags$hr(style = "border-color: #e6f2ff; border-width: 3px; width: 90%; border-radius: 10px; margin: 55px 0 55px 5%;")
               ),
               
               # ----------------- AL FASHER -----------------
               div(
                 id = "al_fasher",
                 class = "hotspot-section",
                 fluidRow(
                   column(
                     6,
                     div(
                       class = "param-box transparent-blue",
                       id = "param_alfasher",
                       style = "display: flex; flex-direction: column; justify-content: center; height: 100%;",
                       div(class = "title", tags$h3(tags$b("Al Fasher (North Darfur)"))),
                       br(), br(),
                       p("After 18 months of siege, during which access to Al Fasher had been completely closed off for both humanitarian and non-humanitarian actors, the RSF took over the city on 26 October 2025. Civilian deaths after the end of the siege are widely underestimated, but they are reported to hover around 2,000 (UN)."),
                       p("The UN says less than half of the 260,000 people estimated to have been in the city before it fell have been accounted for. Moreover, since 26 October, UNICEF states that over 80,000 people have been displaced, with some less reliable sources citing movements of over 100,000 people (AP News). This essentially means that the city remains mostly deserted. Moreover, Zamzam camp, the largest IDP camp located about 15 km south of Al Fasher (700,000 IDPs), has also been attacked several times by the RSF, with major incidents in April 2025 which displaced about 350,000 people. Severe displacement and civilian casualties have been recorded over the past year, raising mortality to 2.5 deaths every 10,000 people per day (MSF)."),
                       p("Most of the displaced population has moved towards Tawila, Melit and Saraf Omra (UNICEF), which are themselves in critical condition. Accordingly, food insecurity is at critical levels, with famine-like conditions reported by both IPC and FSLA in Al Fasher, Tawila, Melit, Um Kadadah, At Tawisha and Al Lait."),
                       p("Physical barriers and conflict limit access to water points. The destruction of sanitation facilities and shelter is widespread, corroborating the scale of damage caused by the ongoing conflict."),
                       p("Most health facilities has been destroyed, with no functioning facilities, while the cholera outbreak has overwhelmed the few facilities that remain. WHO reports that the last remaining hospital in Al Fasher, Saudi Maternity Hospital, has been destroyed after the RSF took over the city."),
                       p("The combination of conflict, cholera, displacement, famine and poor living conditions indicates a severely deteriorated situation, leading to excess mortality.")
                     )
                   ),
                   column(
                     6,
                     div(
                       style = "display: flex; align-items: flex-start; height: 100%; padding-top: 270px; padding-left: 20px",
                       img(src = "alfasher.png", style = "width: 110%; height: 110%")
                     )
                   )
                 ),
                 tags$hr(style = "border-color: #e6f2ff; border-width: 3px; width: 90%; border-radius: 10px; margin: 40px 0 40px 5%;")
               ),
               
               # ----------------- UM DAJ AHMED & BARA -----------------
               div(
                 id = "umdaj_bara",
                 class = "hotspot-section",
                 fluidRow(
                   column(
                     6,
                     div(
                       style = "display: flex; align-items: flex-start; height: 90%; padding-top: 80px; padding-left: 20px",
                       img(src = "bara.png", style = "width: 100%; height: 100%;")
                     )
                   ),
                   column(
                     6,
                     style = "padding-top:30px;",
                     div(class = "title", tags$h3(tags$b("Um Daj Ahmed and Bara (North Kordofan)"))),
                     br(), br(),
                     p("The RSF took over the city on 27 October, prompting the displacement of more than 22,000 people to nearby localities according to IOM, and a total of around 36,000 people across neighbouring Kordofan localities according to UNICEF. Several local news outlets also report evidence of civilians being summarily executed during this period. The SAF subsequently retook the city on 15 November."),
                     p("A similar pattern appears to hold in Bara, which was also taken over by the RSF in October, with thousands of people reportedly displaced according to IOM."),
                     p("Most available data on these localities are likely outdated. Renewed monitoring and needs assessments are required to understand the current situation and inform an appropriate humanitarian response.")
                   )
                 ),
                 br(),
                 tags$hr(style = "border-color: #e6f2ff; border-width: 3px; width: 90%; border-radius: 10px; margin: 55px 0 55px 5%;")
               ),
               
               # ----------------- KHARTOUM & HEALTH COLLAPSE -----------------
               div(
                 id = "khartoum_health",
                 class = "hotspot-section",
                   fluidRow(
                     style = "display: flex; justify-content: center; align-items: center; height: 80%;",
                     column(
                       width = 6,
                       div(class = "title", tags$h3(tags$b("Khartum, Jebel Awlia and the Health System Collapse"))),
                       div(
                         class = "description",
                         br(),
                         p("More than 1 million people have returned to Khartoum over the past year. Homes are damaged and basic services are barely functioning. It is estimated that by late October 2023 about 33,000 buildings had been destroyed in Khartoum (ACAPS). Even though reconstruction is ongoing, it will take several years to re-establish fair living conditions. Many returnees are settling in damaged homes or collective centres, with limited access to clean water, health services, or protection (IOM). HeRAMS reports that 38% of health facilities in Sudan are non-functional and only 14% of hospitals remain operational. Before the conflict, Khartoum used to provide about 70% of the country’s health services, meaning that the capacity of the country to absorb the high number of wounded civilians and military personnel, as well as the high incidence of cholera (which in Khartoum increased ninefold in a matter of 10 days in May 2025 – from 90 to 815 cases), dengue and malaria, is severely compromised. As highlighted by the main drivers map of excess mortality, the health system is the main driver of excess mortality in most parts of the country, even those where the humanitarian situation is non-critical. Especially in the northern, central and eastern parts of the country, cholera incidence and fatality rates are at their highest."),
                         p("Not only health facilities have been affected by the recent conflict. Most water plants have been destroyed, and recent flooding in the region has further compromised access to water."),
                         p("Moreover, Jebel Awlia is considered at risk of famine, and UNICEF states that about one third of the 307,000 children across Khartoum and Jebel Awlia are acutely malnourished as of May 2025."),
                         p("The situation is critical across Khartoum, and especially Jebel Awlia. At a total collapse of the health system across the country is supported by strong evidence.")
                       )
                     ),
                     column(
                       width = 6,
                       div(
                         style = "display: flex; align-items: flex-start; height: 90%; padding-top: 50px; padding-left: 40px",
                         img(src = "khartoum.png", height = "120%", width = "120%")
                       ),
                       div(
                         style = "display: flex; align-items: flex-start; height: 90%; padding-top: 50px; padding-left: 20px",
                         img(src = "health.png", height = "140%", width = "140%")
                       )
                     )
                   ),
                   br(),
                   tags$hr(style = "border-color: #e6f2ff; border-width: 3px; width: 90%; border-radius: 10px; margin: 40px 0 40px 5%;")
               ),
               
               # ----------------- WESTERN BLUE NILE -----------------
               div(
                 id = "western_bluenile",
                 class = "hotspot-section",
                 fluidRow(
                     column(
                       6,
                       div(
                         style = "display: flex; align-items: flex-start; height: 90%; padding-top: 120px; padding-left: 20px",
                         img(src = "bluenile.png", style = "width: 100%; height: 100%;")
                       )
                     ),
                     column(
                       6,
                       style = "padding-top:30px;",
                       div(class = "title", tags$h3(tags$b("Western Blue Nile"))),
                       br(), br(),
                       p("Several overarching crises have been reported in the western part of Blue Nile. Alt Tadamon and Baw are identified as chronic hotspots (while Gaisan and Al Kurmuk are “at-risk” locations) by the Protection Cluster due to the high influx of IDPs and poor living conditions. Most of the population already reports living in inadequate shelter and enduring highly unhygienic sanitation facilities."),
                       p("Unhygienic conditions, coupled with poor water services, are likely to lead to the spread of water-borne diseases. In fact, high cholera incidence and long travel times to the nearest health facility increase the deadliness of the disease."),
                       p("However, the strongest indications of a severe risk of excess mortality come from extreme levels of food insecurity. MSNA data indicates strong food deprivation, especially in Al Kurmuk and Baw, where 80% or more of households report critically low levels of food consumption. The IPC corroborates these findings by assigning an IPC 4 classification to all these localities. The FSLA does not indicate these localities as “at risk of famine”, although it reports that Al Kurmuk has the second highest number of people in need in the country."),
                       p("Overall, western Blue Nile seems to have all systems (food, water, health, and living conditions) in a critical state of failure.")
                     ),
                     br()
                   ),
                   tags$hr(style = "border-color: #e6f2ff; border-width: 3px; width: 90%; border-radius: 10px; margin: 55px 0 55px 5%;")
                 ),
               
               # ----------------- BURAM & AS SUNTA -----------------
               div(
                 id = "buram_assunta",
                 class = "hotspot-section",
                 fluidRow(
                     column(
                       6,
                       div(
                         class = "param-box transparent-blue",
                         id = "param_buram",
                         style = "display: flex; flex-direction: column; justify-content: center; height: 100%;",
                         div(class = "title", tags$h3(tags$b("Buram, As Sunta (South Darfur)"))),
                         br(), br(),
                         p("Poor food outcomes for many IDPs, exhausted coping strategies, and unstable income create a high risk of famine, which is confirmed by both the IPC and FSLA."),
                         p("The protection cluster indicates that both localities are at risk, with more than 42.000 IDPs reported in Buram and 25.000 in As Sunta.  The concerning aspect to these IDP flows is that the influx is growing  steadily over time. This is likely to impact an already fragile shelter system, exposing many households to unsafe, overcrowded, and unsanitary living conditions."),
                         p("Primary health centres nominally function but are practically inaccessible according to HeRams, while strong cholera transmission, high illness incidence in adults and children, and mortality rates just below critical thresholds indicate a severely deteriorated health outcomes."),
                         p("Taken together, severe food insecurity, disease burden, inadequate shelter, and inaccessible health care point to a strong risk of excess mortality, likely to worsen as refugee and IDP arrivals continue.")
                       )
                     ),
                     column(
                       6,
                       div(
                         style = "display: flex; align-items: flex-start; height: 100%; padding-top: 140px; padding-left: 20px",
                         img(src = "southdar.png", style = "width: 100%; height: 100%")
                       )
                     ),
                     br()
                   ),
                   tags$hr(style = "border-color: #e6f2ff; border-width: 3px; width: 90%; border-radius: 10px; margin: 55px 0 55px 5%;")
                 ),
               
               # ----------------- UM DUKHUN -----------------
               div(
                 id = "um_dukhun",
                 class = "hotspot-section",
                 fluidRow(
                     column(
                       6,
                       div(
                         style = "display: flex; align-items: flex-start; height: 90%; padding-top: 140px; padding-left: 20px",
                         img(src = "umduk.png", style = "width: 100%; height: 100%;")
                       )
                     ),
                     column(
                       6,
                       style = "padding-top:30px;",
                       div(class = "title", tags$h3(tags$b("Um Dukhun (Central Darfur)"))),
                       br(), br(),
                       p("In Um Dukhun there are destructive levels of food insecurity, with many pockets of the population struggling to access sufficient food. This is confirmed by the IPC4 rating and by more than 50% of people in CARI P4 (MSNA data). SMART surveys capture that about 1 out of 4 children is considered malnourished by WHZ."),
                       p("While water services are severely compromised, affecting vulnerable populations, shelter and sanitation remain inadequate, especially with the influx of IDPs, leading to extremely poor living conditions. "),
                       p("Access to health facilities is hindered by long travel times, high cholera incidence, and a failing health system, which is further stressed by the large number of people in need. Accordingly, about 60% of children reported having had health issues in the 2 weeks prior to survey which strong incidence of watery diarrhoea and acute respiratory infections (SMART)."),
                       p("The lack of shelter and the spread of unsanitary conditions are contributing to the worsening health outcomes, particularly in vulnerable groups like the elderly and disabled. "),
                       p("This multisystemic breakdown is driving a high risk of excess mortality, particularly exacerbated by high cholera incidence and fatality rates, pointing to a critical situation. This is the only locality for which mortality data is available where there is more than 1 death every 10k people per day indicating that this multisystemic failure already has destructive impacts on the population.")
                     ),
                     br()
                   ),
                   tags$hr(style = "border-color: #e6f2ff; border-width: 3px; width: 90%; border-radius: 10px; margin: 55px 0 55px 5%;")
                 ),
               
               # ----------------- HABIL A (WEST DARFUR) -----------------
               div(
                 id = "habila",
                 class = "hotspot-section",
                 fluidRow(
                     column(
                       6,
                       div(
                         class = "param-box transparent-blue",
                         id = "param_habila",
                         style = "display: flex; flex-direction: column; justify-content: center; height: 100%;",
                         div(class = "title", tags$h3(tags$b("Habila (West Darfur)"))),
                         br(), br(),
                         p("Most households struggle to access enough food as highlighted by the IPC 4 classification. However, functioning markets and unexhausted coping strategies do not indicate the plausibility for further deterioration. "),
                         p("Moreover, the severe failure of the WASH and health system do seem to have devastating impact on the population. The water system has critically failed: unimproved sources are widely used, households reduce consumption, and long queues and poor access are widespread. High cholera incidence and fatality rates, together with widespread use of unsafe water, clearly signal very poor health outcomes."),
                         p("Living conditions are systemically deficient, with unimproved, unhygienic shelter and sanitation exposing people to the elements and facilitating disease transmission. "),
                         p("Health facilities are effectively non-functional and largely unreachable, while strong illness incidence and intense cholera transmission further undermine any remaining capacity. "),
                         p("Overall, access to health care, safe water, and adequate living conditions is severely compromised, and the situation is critical with a substantially elevated risk of excess mortality.")
                       )
                     ),
                     column(
                       6,
                       div(
                         style = "display: flex; align-items: flex-start; height: 100%; padding-top: 120px; padding-left: 20px",
                         img(src = "habilaw.png", style = "width: 100%; height: 100%")
                       )
                     ),
                     br()
                   ),
                   tags$hr(style = "border-color: #e6f2ff; border-width: 3px; width: 90%; border-radius: 10px; margin: 55px 0 55px 5%;")
                 ),
               
               # ----------------- REIFI TELKOK -----------------
               div(
                 id = "reifi_telkok",
                 class = "hotspot-section",
                 fluidRow(
                     column(
                       6,
                       div(
                         style = "display: flex; align-items: flex-start; height: 90%; padding-top: 120px; padding-left: 20px",
                         img(src = "telkok.png", style = "width: 100%; height: 100%;")
                       )
                     ),
                     column(
                       6,
                       style = "padding-top:30px;",
                       div(class = "title", tags$h3(tags$b("Reifi Telkok (Kassala)"))),
                       br(), br(),
                       p("Food security currently appears stable, with no reported food insecurity and coping strategies still intact. "),
                       p("High water insecurity is a major concern: many households rely on heavy consumption-based coping and must spend more than 30 minutes reaching water sources. This has also been reported in many localities in the region (both Kassala and Gedaref)."),
                       p("Living conditions are critically compromised, as widespread hygiene problems, poor shelter, and unimproved sanitation pose serious risks to most of the population"),
                       p("Health facilities function and are geographically accessible, but the heavy cholera caseload likely exceeds their effective capacity. High cholera incidence and fatality rates signal severe, unmet health needs."),
                       p("Overall, intense cholera transmission combined with degraded WASH and shelter conditions creates a severe hazard, and further deterioration in health outcomes is highly likely in the near to medium term. More data needs to be made available on this locality as it has not been taken into account by other humanitarian reports and organizations. ")
                     ),
                     br()
                   ),
                   tags$hr(style = "border-color: #e6f2ff; border-width: 3px; width: 90%; border-radius: 10px; margin: 55px 0 55px 5%;")
               )
             )
           )
  )
  ),
############################ CONTEXT MAPS ###################################
tabPanel(
  "Context Maps",
  div(
    class = "content-wrapper",
    style = "margin-top: 70px;",
    tags$head(includeCSS("styles.css")),
    
    div(
      class = "hotspots-layout",
      
      # LEFT SIDE PANEL (same sidebar, same links / ids)
      div(
        style = "margin-top: 180px;",
        class = "hotspots-sidebar",
        div(
          class = "hotspots-sidebar-box",
          div(class = "hotspots-nav-title", "Contents"),
          tags$ul(
            class = "hotspots-nav-links",
            tags$li(a(href = "#areasinfluence","Areas of Influence")),
            tags$li(a(href = "#dtm",           "Displacement Tracking Matrix")),
            tags$li(a(href = "#refugee",       "Refugee Movements")),
            tags$li(a(href = "#returnee",      "Returnee Movements")),
            tags$li(a(href = "#protection",    "Protection Hotspots")),
            tags$li(a(href = "#access",        "Humanitarian Access")),
            tags$li(a(href = "#ipc",           "IPC Classification")),
            tags$li(a(href = "#food",          "Food System Evaluation")),
            tags$li(a(href = "#healthinf",      "Health Infrastructure")),
            tags$li(a(href = "#healthsystem",   "Health System Evaluation")),
            tags$li(a(href = "#healthoutcomes","Health Outcomes Evaluation")),
            tags$li(a(href = "#water",         "Water System Evaluation")),
            tags$li(a(href = "#livingcond",    "Living Conditions Evaluation")),
          )
        )
      ),
      
      # RIGHT SIDE: MAPS ONLY (title + map)
      div(
        class = "hotspots-main",
        style = "margin-top: 0px; margin-left: 0px;",
        
        tags$h2(tags$b("Contextual Maps")),
        
        br(),
        div(
          id = "areasinfluence",
          class = "hotspot-section",
          fluidRow(
            column(
              width = 12,
              div(
                class = "title",
                tags$h3(tags$b("Areas of Influence (Liveuamap) - November 2025"))
              ),
              div(
                style = "display: flex; justify-content: center; padding-top: 20px;",
                img(src = "areainfluence.png", style = "width: 80%; height: auto;"),

              ),
              br(),
              div(
                style = "font-size: 85%; color: #666; text-align: center;",
                "Source: ",
                a("Liveuamap", href = "https://sudan.liveuamap.com/", target = "_blank", style = "margin-left: 5px; margin-top: 20px")
              )
            )
          ),
          tags$hr(style = "border-color: #e6f2ff; border-width: 3px;
                           width: 90%; border-radius: 10px;
                           margin: 40px 0 40px 5%;")
        ),
        

        div(
          id = "dtm",
          class = "hotspot-section",
          fluidRow(
            column(
              width = 12,
              div(
                class = "title",
                tags$h3(tags$b("Displacement Tracking Matrix (IOM) - October 2025"))
              ),
              div(
                style = "display: flex; justify-content: center; padding-top: 20px;",
                img(src = "dtm.png", style = "width: 80%; height: auto;")
              ),
              br(),
              div(
                style = "font-size: 85%; color: #666; text-align: center;",
                "Source: ",
                a("IOM - DTM", href = "https://dtm.iom.int/sites/g/files/tmzbdl1461/files/reports/DTM%20Sudan%20Mobility%20Update%2022_0.pdf", target = "_blank", style = "margin-left: 5px; margin-top: 20px")
              )
            )
          ),
          tags$hr(style = "border-color: #e6f2ff; border-width: 3px;
                           width: 90%; border-radius: 10px;
                           margin: 40px 0 40px 5%;")
        ),
        
        div(
          id = "refugee",
          class = "hotspot-section",
          fluidRow(
            column(
              width = 12,
              div(
                class = "title",
                tags$h3(tags$b("Refugee Movements (UNHCR) - November 2025"))
              ),
              div(
                style = "display: flex; justify-content: center; padding-top: 20px;",
                img(src = "refugee.png", style = "width: 80%; height: auto;")
              ),
              br(),
              div(
                style = "font-size: 85%; color: #666; text-align: center;",
                "Source: ",
                a("UNHCR", href = "https://data.unhcr.org/en/documents/details/119860", target = "_blank", style = "margin-left: 5px; margin-top: 20px")
              )
            )
          ),
          tags$hr(style = "border-color: #e6f2ff; border-width: 3px;
                           width: 90%; border-radius: 10px;
                           margin: 40px 0 40px 5%;")
        ),
        
        
        div(
          id = "returnee",
          class = "hotspot-section",
          fluidRow(
            column(
              width = 12,
              div(
                class = "title",
                tags$h3(tags$b("Returnee Movements (UNHCR) - November 2025"))
              ),
              div(
                style = "display: flex; justify-content: center; padding-top: 20px;",
                img(src = "returnee2.png", style = "width: 80%; height: auto;")
              ),
              br(),
              div(
                style = "font-size: 85%; color: #666; text-align: center;",
                "Source: ",
                a("UNHCR", href = "https://reliefweb.int/map/south-sudan/south-sudan-refugees-returnees-and-idp-map-31-october-2025", target = "_blank", style = "margin-left: 5px; margin-top: 20px")
              )
            )
          ),
          tags$hr(style = "border-color: #e6f2ff; border-width: 3px;
                           width: 90%; border-radius: 10px;
                           margin: 40px 0 40px 5%;")
        ),
        

        div(
          id = "protection",
          class = "hotspot-section",
          fluidRow(
            column(
              width = 12,
              div(
                class = "title",
                tags$h3(tags$b("Protection Hotspots (Protection Cluster) - August 2025"))
              ),
              div(
                style = "display: flex; justify-content: center; padding-top: 20px;",
                img(src = "protection.png", style = "width: 80%; height: auto;")
              ),
              br(),
              div(
                style = "font-size: 85%; color: #666; text-align: center;",
                "Source: ",
                a("UNHCR", href = "https://reliefweb.int/report/sudan/protection-hotspots-sudan-31-august-2025", target = "_blank", style = "margin-left: 5px; margin-top: 20px")
              )
            )
          ),
          tags$hr(style = "border-color: #e6f2ff; border-width: 3px;
                           width: 90%; border-radius: 10px;
                           margin: 40px 0 40px 5%;")
        ),
        
        div(
          id = "access",
          class = "hotspot-section",
          fluidRow(
            column(
              width = 12,
              div(
                class = "title",
                tags$h3(tags$b("Humanitarian Access Map (OCHA) - August 2025"))
              ),
              div(
                style = "display: flex; justify-content: center; padding-top: 20px;",
                img(src = "accessmap.png", style = "width: 80%; height: 80%;")
              ),
              br(),
              div(
                style = "font-size: 85%; color: #666; text-align: center;",
                "Source: ",
                a("UN OCHA", href = "https://www.unocha.org/publications/report/sudan/sudan-humanitarian-access-snapshot-august-2025", target = "_blank", style = "margin-left: 5px; margin-top: 20px")
              )
            )
          ),
          tags$hr(style = "border-color: #e6f2ff; border-width: 3px;
                           width: 90%; border-radius: 10px;
                           margin: 40px 0 40px 5%;")
        ),
        
        
        div(
          id = "ipc",
          class = "hotspot-section",
          fluidRow(
            column(
              width = 12,
              div(
                class = "title",
                tags$h3(tags$b("IPC Classification Map (IPC)- November 2025"))
              ),
              div(
                style = "display: flex; justify-content: center; padding-top: 20px;",
                img(src = "ipc.png", style = "width: 80%; height: auto;")
              ),
              br(),
              div(
                style = "font-size: 85%; color: #666; text-align: center;",
                "Source: ",
                a("IPC", href = "https://www.ipcinfo.org/fileadmin/user_upload/ipcinfo/docs/IPC_Sudan_Acute_Food_Insecurity_Sep2025_May2026_Special_Snapshot.pdf", target = "_blank", style = "margin-left: 5px; margin-top: 20px")
              )
            )

          ),
          tags$hr(style = "border-color: #e6f2ff; border-width: 3px;
                           width: 90%; border-radius: 10px;
                           margin: 40px 0 40px 5%;")
        ),
        

        div(
          id = "food",
          class = "hotspot-section",
          fluidRow(
            column(
              width = 12,
              div(
                class = "title",
                tags$h3(tags$b("Food System Evaluation (REACH) - September 2025"))
              ),     
              
              div(
                style = "display: flex; justify-content: center; gap: 20px; padding-top: 20px;",
                img(src = "food.png", style = "width: 70%; height: auto;")
              ),
              br(),
              div(
                style = "font-size: 85%; color: #666; text-align: center;",
                "Source: REACH (consult `Data Download`)"
                
              )
            )
          ),
          tags$hr(style = "border-color: #e6f2ff; border-width: 3px;
                           width: 90%; border-radius: 10px;
                           margin: 40px 0 40px 5%;")
        ),
        

        div(
          id = "healthinf",
          class = "hotspot-section",
          fluidRow(
            column(
              width = 12,
              div(
                class = "title",
                tags$h3(tags$b("Health Infrastructure (HeRAMS) - 2025"))
              ),
              div(
                style = "display: flex; justify-content: center; padding-top: 20px;",
                img(src = "healthinfrastructure.png", style = "width: 80%; height: auto;")
              ),
              br(),
              div(
                style = "font-size: 85%; color: #666; text-align: center;",
                "Source: ",
                a("HeRAMS", href = "https://www.who.int/publications/m/item/herams-sudan-baseline-report-2025-operational-status-of-the-health-system",
                  target = "_blank", style = "margin-left: 5px; margin-top: 20px")
              )
              
            )
          ),
          tags$hr(style = "border-color: #e6f2ff; border-width: 3px;
                           width: 90%; border-radius: 10px;
                           margin: 40px 0 40px 5%;")
        ),
        
        
        div(
          id = "healthsystem",
          class = "hotspot-section",
          fluidRow(
            column(
              width = 12,
              div(
                class = "title",
                tags$h3(tags$b("Health System Evaluation (REACH) - September 2025"))
              ),
              div(
                style = "display: flex; justify-content: center; padding-top: 20px;",
                img(src = "healthsyst.png", style = "width: 80%; height: auto;")
              ),
              br(),
              div(
                style = "font-size: 85%; color: #666; text-align: center;",
                "Source: REACH (consult `Data Download`)"

              )
            )
          ),
          tags$hr(style = "border-color: #e6f2ff; border-width: 3px;
                           width: 90%; border-radius: 10px;
                           margin: 40px 0 40px 5%;")
        ),
        
        div(
          id = "healthoutcomes",
          class = "hotspot-section",
          fluidRow(
            column(
              width = 12,
              div(
                class = "title",
                tags$h3(tags$b("Health Outcomes Evaluation (REACH) - September 2025"))
              ),
              div(
                style = "display: flex; justify-content: center; padding-top: 20px;",
                img(src = "houtcomes.png", style = "width: 80%; height: auto;")
              ),
              br(),
              div(
                style = "font-size: 85%; color: #666; text-align: center;",
                "Source: REACH (consult `Data Download`)"
                
              )
            )
          ),
          tags$hr(style = "border-color: #e6f2ff; border-width: 3px;
                           width: 90%; border-radius: 10px;
                           margin: 40px 0 40px 5%;")
        ),
        
        div(
          id = "water",
          class = "hotspot-section",
          fluidRow(
            column(
              width = 12,
              div(
                class = "title",
                tags$h3(tags$b("Water System Evaluation (REACH) - September 2025"))
              ),
              div(
                style = "display: flex; justify-content: center; padding-top: 20px;",
                img(src = "water.png", style = "width: 80%; height: auto;")
              ),
              br(),
              div(
                style = "font-size: 85%; color: #666; text-align: center;",
                "Source: REACH (consult `Data Download`)"
                
              )
            )
          ),
          tags$hr(style = "border-color: #e6f2ff; border-width: 3px;
                           width: 90%; border-radius: 10px;
                           margin: 40px 0 40px 5%;")
        ),
        
        div(
          id = "livingcond",
          class = "hotspot-section",
          fluidRow(
            column(
              width = 12,
              div(
                class = "title",
                tags$h3(tags$b("Living Conditions Evaluation (REACH) - September 2025"))
              ),
              div(
                style = "display: flex; justify-content: center; padding-top: 20px;",
                img(src = "living.png", style = "width: 80%; height: auto;")
              ),
              br(),
              div(
                style = "font-size: 85%; color: #666; text-align: center;",
                "Source: REACH (consult `Data Download`)"
                
              )
            )
          ),
          tags$hr(style = "border-color: #e6f2ff; border-width: 3px;
                           width: 90%; border-radius: 10px;
                           margin: 40px 0 40px 5%;")
        )
        

    
      )
    )
  )
),
  
  
  ############################ DATA DOWNLOAD ###################################
  tabPanel(
    "Data Download",
    value = "data_download",
    div( 
      class = "content-wrapper",
      style = "margin-top: 70px;",
      
      tags$h2(tags$b("Download Source Data")),
      
      p("Click the button below to download the file with the full data used to conduct the analysis. Please mind that qualitative data that was used is not referenced."),
      br(),
      
      downloadButton(
        outputId = "download_ana_file",
        label = "Download Source Data",
        style = "
    border-radius: 25px;
    padding: 10px 20px;
    background-color: #000000;
    color: #ffffff;
    border-color: #000000;
  "
      ),
      
      br(), br(),    br(),
      
      h4(tags$b("Preview of first sheet")),
      br(),
      p("Table below shows the contents of the first sheet of the file."),
      
      br(),   br(),
      
      DT::dataTableOutput("ana_preview")
    )
  )
  


# # Bottom-right logo + socials
# div(
#   style = "position: fixed; bottom: 16px; right: 8px;
#            display: flex; flex-direction: column;
#            align-items: center; justify-content: flex-end; gap: 6px;",
#   
#   a(
#     href   = "https://www.impact-initiatives.org",
#     target = "_blank",
#     class  = "impact",
#     style  = "display: inline-block; margin-bottom: 70px;",
#     tags$img(
#       src   = "logo.png",
#       style = "width: 160px; height: auto; display: block;
#                transform: rotate(-90deg); transform-origin: center center;"
#     )
#   ),
#   
#   a(
#     href   = "https://x.com/impact_init?lang=en",
#     target = "_blank",
#     class  = "twitter",
#     style  = "color: black; font-size: 30px;",
#     HTML('<i class="fab fa-twitter"></i>')
#   ),
#   
#   a(
#     href   = "https://www.linkedin.com/company/impact-initiatives",
#     target = "_blank",
#     class  = "linkedin",
#     style  = "color: black; font-size: 30px;",
#     HTML('<i class="fab fa-linkedin"></i>')
#   )
# )


)
)

# =============================== SERVER =======================================
server <- function(input, output, session) {
  
  # Auto-focus select box once UI is ready
  session$onFlushed(function() {
    session$sendCustomMessage("focusLocSelect", list())
  }, once = TRUE)
  
  
  
  ##################################### 1. LEFT SIDE PANEL #####################
  
  observeEvent(input$toggle_map_desc, {
    if (input$toggle_map_desc) {
      shinyjs::show("map_description_box")
    } else {
      shinyjs::hide("map_description_box")
    }
  })
  output$map_type_text <- renderUI({
    req(input$map_type)
    
    if (input$map_type == "deepdive") {
      div(
        class = "help",
        tags$p("This map displays data on the Risk of Excess Mortality (RoEM) based on the Acute Needs Analysis Framework (ANA). Localities are assigned to one of five levels of severity."),
        tags$p(tags$strong("• Acute Needs (AN):"), " High severity gaps in key mortality drivers, which increase RoEM risk when they co-occur."),
        tags$p(tags$strong("• AN (Suspect RoEM):"), " Severe gaps indicate high likelihood of developing RoEM in the medium to long term."),
        tags$p(tags$strong("• Risk of Excess Mortality (RoEM):"), " Life-threatening needs are increasing the risk of mortality based on available data."),
        tags$p(tags$strong("• RoEM (Suspect EM):"), " A situation where excess mortality is suspected but not confirmed."),
        tags$p(tags$strong("• Excess Mortality:"), " Mortality exceeding baseline levels, expressed as number of excess deaths above 1 every 10k individuals per day.")
      )
      
    } else if (input$map_type == "flat") {
      div(
        class = "help",
        tags$p("This map displays simplified levels of the Risk of Excess Mortality (RoEM), grouped into 3 categories."),
        tags$p(tags$strong("• Acute Needs (AN):"), " High severity gaps in key mortality drivers."),
        tags$p(tags$strong("• Risk of Excess Mortality (RoEM):"), " Life-threatening needs are increasing the risk of mortality."),
        tags$p(tags$strong("• Excess Mortality:"), " Mortality exceeding baseline levels (>1/10k/day).")
      )
      
    } else if (input$map_type == "driver") {
      div(
        class = "help",
        tags$p("This map shows the main driver contributing to Risk of Excess Mortality (RoEM) in each locality."),
        tags$p(tags$strong("• Health system:"), " Functionality and access to healthcare and nutrition."),
        tags$p(tags$strong("• Food system:"), " Food availability, access, consumption, and market constraints."),
        tags$p(tags$strong("• Water system:"), " Access to clean water and adequacy of sanitation."),
        tags$p(tags$strong("• Living conditions:"), " Housing, environment, and exposure risks."),
        tags$p(tags$strong("• Conflict:"), " Insecurity, violence, access disruption."),
        tags$p(tags$strong("• Displacement:"), " Forced movement and its impacts on service access.")
      )
    }
  })
  
    
  
  
  
  ################################## 2. MAP ####################################
  # Base leaflet
  output$map <- renderLeaflet({
    leaflet::leaflet(
      options = leaflet::leafletOptions(
        worldCopyJump      = TRUE,
        preferCanvas       = TRUE,
        zoomControl        = F,
        attributionControl = FALSE,
        scrollWheelZoom    = TRUE
      )
    ) |>
      leaflet::addProviderTiles(leaflet::providers$CartoDB.Positron) |>
      leaflet::setView(lng = 30, lat = 13, zoom = 6) |>
      htmlwidgets::onRender("
        function(el,x){
          var map = this;
          if (HTMLWidgets && HTMLWidgets.find) {
            var w = HTMLWidgets.find(el.id);
            if (w && w.getMap) {
              map = w.getMap();
            }
          }
          window._anaMap = map;
          if (map.zoomControl) map.zoomControl.setPosition('topright');
        }
      ")
  })
  
  # Helper: base polygons
  draw_base_polygons <- function(map_type) {
    if (map_type == "deepdive") {
      map_dat <- summary_data %>%
        mutate(
          label_val = if_else(
            is.na(deep_dive),
            "No data",
            as.character(deep_dive)
          )
        )
        
      leafletProxy("map") %>%
        clearGroup("base") %>%
        clearControls() %>%
        addPolygons(
          data = map_dat,
          group = "base",
          layerId = ~ADM2_PCODE,
          fillColor = ~pal_deepdive(deep_dive),  # Apply the color factor correctly
          color = "#555555",
          weight = 0.7,
          opacity = 1,
          fillOpacity = 0.7,
          label = ~paste0(ADM2_EN, " — ", label_val),
          highlightOptions = highlightOptions(
            weight = 1.4,
            color = "#222222",
            opacity = 1,
            bringToFront = TRUE
          )
        ) %>%
        addLegend(
          position = "bottomleft",
          colors = unname(driver_colors_roem),
          labels = names(driver_colors_roem),
          title = "Acute Needs Categorization (5-Level)",
          opacity = 0.9
        )
    }
    else if (map_type == "driver") {
      map_dat <- summary_data %>%
        mutate(
          label_val = if_else(
            is.na(main_driver),
            "No data",
            as.character(main_driver)
          )
        )
      
      leafletProxy("map") %>%
        clearGroup("base") %>%
        clearControls() %>%
        addPolygons(
          data        = map_dat,
          group       = "base",
          layerId     = ~ADM2_PCODE,
          fillColor   = ~pal_driver(main_driver),
          color       = "#555555",
          weight      = 0.7,
          opacity     = 1,
          fillOpacity = 0.7,
          label       = ~paste0(ADM2_EN, " — ", label_val),
          highlightOptions = highlightOptions(
            weight = 1.4,
            color  = "#222222",
            opacity = 1,
            bringToFront = TRUE
          )
        ) %>%
        addLegend(
          position = "bottomleft",
          colors   = unname(driver_colors),
          labels   = names(driver_colors),
          title    = "Main driver of RoEM",
          opacity  = 0.9
        )
    }  else if (map_type == "flat") {
      map_dat <- summary_data %>%
        mutate(
          label_val = if_else(
            is.na(deep_dive_flat),
            "No data",
            as.character(deep_dive_flat)
          )
        )
      
      leafletProxy("map") %>%
        clearGroup("base") %>%
        clearControls() %>%
        addPolygons(
          data        = map_dat,
          group       = "base",
          layerId     = ~ADM2_PCODE,
          fillColor   = ~pal_deepdive_flat(deep_dive_flat),
          color       = "#555555",
          weight      = 0.7,
          opacity     = 1,
          fillOpacity = 0.7,
          label       = ~paste0(ADM2_EN, " — ", label_val),
          highlightOptions = highlightOptions(
            weight = 1.4,
            color  = "#222222",
            opacity = 1,
            bringToFront = TRUE
          )
        ) %>%
        addLegend(
          position = "bottomleft",
          colors   = unname(driver_colors_roem_flat),
          labels   = names(driver_colors_roem_flat),
          title    = "Acute Needs Categorization (3-Level)",
          opacity  = 0.9
        )
    } 
  }
  
  # Helper: selection overlay + connector JS
  draw_selection_overlay <- function(map_type, loc_name) {
    leafletProxy("map") %>% clearGroup("selected")
    
    if (!nzchar(loc_name)) {
      session$sendCustomMessage("updateConnector", list(active = FALSE))
      return(invisible(NULL))
    }
    
    if (map_type == "deepdive") {
      map_dat <- summary_data %>%
        mutate(
          label_val = if_else(
            is.na(deep_dive),
            "No data",
            as.character(deep_dive)
          )
        )
      sel <- map_dat[map_dat$ADM2_EN == loc_name, , drop = FALSE]
      if (nrow(sel) == 0) {
        session$sendCustomMessage("updateConnector", list(active = FALSE))
        return(invisible(NULL))
      }
      col <- pal_deepdive(sel$deep_dive)
    } else if (map_type == "driver") {
      map_dat <- summary_data %>%
        mutate(
          label_val = if_else(
            is.na(main_driver),
            "No data",
            as.character(main_driver)
          )
        )
      sel <- map_dat[map_dat$ADM2_EN == loc_name, , drop = FALSE]
      if (nrow(sel) == 0) {
        session$sendCustomMessage("updateConnector", list(active = FALSE))
        return(invisible(NULL))
      }
      col <- pal_driver(sel$main_driver)
    } else if (map_type == "flat") {
      map_dat <- summary_data %>%
        mutate(
          label_val = if_else(
            is.na(deep_dive_flat),
            "No data",
            as.character(deep_dive_flat)
          )
        )
      sel <- map_dat[map_dat$ADM2_EN == loc_name, , drop = FALSE]
      if (nrow(sel) == 0) {
        session$sendCustomMessage("updateConnector", list(active = FALSE))
        return(invisible(NULL))
      }
      col <- pal_deepdive_flat(sel$deep_dive_flat)
    }
    
    leafletProxy("map") %>%
      addPolygons(
        data        = sel,
        group       = "selected",
        fillColor   = col,
        fillOpacity = 0.9,
        color       = "#222222",
        weight      = 1.8,
        opacity     = 1
      )
    
    cent <- summary_centroids[summary_centroids$ADM2_EN == loc_name, , drop = FALSE]
    if (nrow(cent) == 0) {
      session$sendCustomMessage("updateConnector", list(active = FALSE))
      return(invisible(NULL))
    }
    
    session$sendCustomMessage(
      "updateConnector",
      list(
        active = TRUE,
        lat    = cent$lat[1],
        lng    = cent$lon[1]
      )
    )
  }
  
  # Base polygons when map type changes
  observeEvent(input$map_type, {
    mt <- input$map_type %||% "deepdive"
    draw_base_polygons(mt)
    draw_selection_overlay(mt, input$loc_select %||% "")
  }, ignoreInit = FALSE)
  
  
  # show/hide panel and update connector on selection change
  observeEvent(input$loc_select, {
    loc <- input$loc_select %||% ""
    
    if (nzchar(loc)) {
      shinyjs::show("info-panel")
    } else {
      shinyjs::hide("info-panel")
    }
    
    mt <- input$map_type %||% "deepdive"
    draw_selection_overlay(mt, loc)
  }, ignoreInit = TRUE)
  
  # Map click: update selectize
  observeEvent(input$map_shape_click, {
    click <- input$map_shape_click
    req(click$id)
    row <- adm2_nog[adm2_nog$ADM2_PCODE == click$id, ]
    if (nrow(row) == 1) {
      updateSelectizeInput(session, "loc_select", selected = row$ADM2_EN)
    }
  })
  
  
  
  #################################### 3. RIGHT SIDE PANEL #####################
  
  # helper: extract one indicator value for a locality
  get_field <- function(loc, ind) {
    out <- ana_long %>%
      filter(`Administrative Level Label` == loc, indicator == ind) %>%
      pull(value)
    if (length(out) == 0) return("—")
    out[1]
  }
  
  # header: locality – State
  output$sidebar_header <- renderUI({
    req(input$loc_select)
    
    state <- ana_long %>%
      filter(`Administrative Level Label` == input$loc_select) %>%
      pull(State)
    state <- ifelse(length(state) == 0, "", unique(state)[1])
    
    div(
      class = "side-title",
      paste0(input$loc_select),
      span(
        paste0(" – ", state),
        class = "fade-state",
        style = "font-size:13px;"
      )
    )
  })
  
  # fields
  output$sidebar_table <- renderUI({
    req(input$loc_select)
    
    sev   <- get_field(input$loc_select, "Severity")
    con   <- get_field(input$loc_select, "Conclusion")
    cons  <- get_field(input$loc_select, "Conclusion Short")
    plaus <- get_field(input$loc_select, "Plausibility judgement")
    main  <- get_field(input$loc_select, "Main driver")
    sec   <- get_field(input$loc_select, "Secondary driver")
    
    tagList(
      div(class = "side-label", "Severity:"),
      div(class = "side-value", sev),
      
      div(class = "side-label", "Conclusion:"),
      div(class = "side-value", con),
      
      div(class = "side-label", "Conclusion summary:"),
      div(class = "side-value", cons),
      
      div(class = "side-label", "Plausibility judgement:"),
      div(class = "side-value", plaus),
      
      div(class = "side-label", "Main driver:"),
      div(class = "side-value", main),
      
      div(class = "side-label", "Secondary driver:"),
      div(class = "side-value", sec)
    )
  })
  
  # small donut for selected locality
  loc_df <- reactive({
    req(input$loc_select)
    deep_shares %>% filter(`Administrative Level Label` == input$loc_select)
  })
  
  output$mini_donut <- renderPlot({
    df <- loc_df()
    req(nrow(df) > 0)
    
    ggplot(df, aes(x = 2, y = share, fill = conclusion)) +
      geom_col(color = "white", width = 1) +
      coord_polar(theta = "y") +
      xlim(0.5, 3) +
      scale_fill_manual(values = conclusion_colors) +
      theme_void() +
      theme(plot.title = element_blank()) +  # Remove the title (if any)
      guides(fill = guide_legend(title = NULL))  # Remove the legend title
  })
  
  # Show list of conclusions by system
  output$within_system_conclusions <- renderUI({
    req(input$loc_select)
    
    # Filter for rows where the value starts with C1, C2, or C3 for the selected locality
    system_conclusions <- ana_long %>%
      filter(`Administrative Level Label` == input$loc_select,
             indicator == "Conclusion",
             value %in% c("C1: Systemic Failure", "C2: Pockets of severity", 
                          "C3: Mitigation/no acute needs", "Inconclusive")) %>%
      arrange(System) %>%  # Ensure 'System' column is referenced correctly
      select(System, value)
    
    # Define conclusion colors
    conclusion_colors <- c(
      "C1: Systemic Failure"          = "#FF9F9A",
      "C2: Pockets of severity"       = "#FFE4A3",
      "C3: Mitigation/no acute needs" = "#C9F2A3",
      "Inconclusive"                  = "#D3D3D3"
    )
    
    # Create a list of outputs with colored dots and system names
    output_list <- lapply(1:nrow(system_conclusions), function(i) {
      system_name <- system_conclusions$System[i]
      conclusion_value <- system_conclusions$value[i]
      color <- conclusion_colors[conclusion_value]
      
      # Create a div with a colored dot and the system name
      div(
        style = paste("display: flex; align-items: center; margin-bottom: 5px;"),
        span(
          style = paste("width: 10px; height: 10px; background-color:", color, "; border-radius: 50%; margin-right: 5px;")
        ),
        system_name
      )
    })
    
    # Return the rendered UI elements
    do.call(tagList, output_list)
  })
  
  loc_df2 <- reactive({
    req(input$loc_select)
    deep_shares2 %>% filter(`Administrative Level Label` == input$loc_select)
  })
  
  
  output$mini_donut2 <- renderPlot({
    df <- loc_df2()
    req(nrow(df) > 0)
    
    ggplot(df, aes(x = 2, y = share, fill = conclusion)) +
      geom_col(color = "white", width = 1) +
      coord_polar(theta = "y") +
      xlim(0.5, 3) +
      scale_fill_manual(values = conclusion_colors2) +
      theme_void() +
      theme(plot.title = element_blank()) +  # Remove the title (if any)
      guides(fill = guide_legend(title = NULL))  # Remove the legend title
  })
  
  
  
  output$across_system_conclusions <- renderUI({
    req(input$loc_select)
    
    # Filter for rows where the value starts with C1, C2, or C3 for the selected locality
    system_conclusions <- ana_long %>%
      filter(`Administrative Level Label` == input$loc_select,
             indicator == "Conclusion",
             value %in% c("C1: Strong Interaction", "C2: Limited or indirect Interaction", 
                          "C3: No interaction")) %>%
      arrange(System) %>%  # Ensure 'System' column is referenced correctly
      select(System, value)
    
    # Define conclusion colors
    conclusion_colors <- c(
      "C1: Strong Interaction"          = "#FF9F9A",
      "C2: Limited or indirect Interaction"       = "#FFE4A3",
      "C3: No interaction" = "#C9F2A3"
    )
    
    # Create a list of outputs with colored dots and system names
    output_list <- lapply(1:nrow(system_conclusions), function(i) {
      system_name <- system_conclusions$System[i]
      conclusion_value <- system_conclusions$value[i]
      color <- conclusion_colors[conclusion_value]
      
      # Create a div with a colored dot and the system name
      div(
        style = paste("display: flex; align-items: center; margin-bottom: 5px;"),
        span(
          style = paste("width: 10px; height: 10px; background-color:", color, "; border-radius: 50%; margin-right: 5px;")
        ),
        system_name
      )
    })
    
    # Return the rendered UI elements
    do.call(tagList, output_list)
  })
  
  
  
  #################################### 4. LOCALITY REPORT MAP  #################
  
  # Create small map for selected locality
  output$small_map <- renderLeaflet({
    req(input$locality_report)
    
    # Filter the data for the selected locality
    locality_data <- summary_data %>%
      filter(ADM2_EN == input$locality_report)
    
    # If no data, return empty map
    if (nrow(locality_data) == 0) {
      return(leaflet() %>% addTiles())
    }
    
    # Get the centroid of the selected locality for centering the map
    centroid <- locality_data %>%
      st_centroid() %>%
      st_coordinates() %>%
      as.data.frame()
    
    # Draw the base polygons
    leaflet(options = leaflet::leafletOptions(
      worldCopyJump = TRUE,
      preferCanvas = TRUE,
      zoomControl = F,
      attributionControl = FALSE,
      scrollWheelZoom = TRUE
    )) %>%
      addProviderTiles(providers$CartoDB.Positron) %>%
      setView(lng = centroid$X, lat = centroid$Y, zoom = 7) %>%
      addPolygons(
        data = locality_data,
        color = "#222222",
        weight = 3,  # Bold border
        fillColor = "#FF6347",  # Example fill color
        fillOpacity = 0.5
      )
  })
  

  output$download_locality_data <- downloadHandler(
    filename = function() {
      paste0(
        "locality_data_",
        input$locality_report, "_",
        Sys.Date(),
        ".xlsx"
      )
    },
    content = function(file) {
      req(input$locality_report)
      
      # 1. Filter ana_long and ana_base_data by selected locality
      locality_data <- ana_long_download %>%
        dplyr::filter(`Administrative Level Label` == input$locality_report)
      
      base_locality_data <- ana_base_data %>%
        dplyr::filter(`Administrative Level Label` == input$locality_report)
      
      # 2. Create workbook and first sheet
      wb <- openxlsx::createWorkbook()
      sheet_name_1 <- "ANA Deep Dive Results"
      openxlsx::addWorksheet(wb, sheetName = sheet_name_1)
      
      # 3. If no data in ana_long: write message in first row and exit
      if (nrow(locality_data) == 0) {
        openxlsx::writeData(
          wb,
          sheet = sheet_name_1,
          x = "NO DATA AVAILABLE FOR THIS LOCALITY",
          startRow = 1,
          startCol = 1
        )
        openxlsx::saveWorkbook(wb, file, overwrite = TRUE)
        return(NULL)
      }
      
      # 4. Write filtered ana_long data to first sheet
      openxlsx::writeData(
        wb,
        sheet = sheet_name_1,
        x = locality_data,
        startRow = 1,
        startCol = 1,
        headerStyle = NULL
      )
      
      # 5. Header style (reused for both sheets)
      header_style <- openxlsx::createStyle(
        textDecoration = "bold",
        halign        = "center",
        valign        = "center",
        wrapText      = TRUE,
        fgFill        = "#e34444",
        fontColour    = "#FFFFFF"
      )
      
      openxlsx::addStyle(
        wb,
        sheet      = sheet_name_1,
        style      = header_style,
        rows       = 1,
        cols       = 1:ncol(locality_data),
        gridExpand = TRUE
      )
      
      # 6. Header row height (sheet 1)
      openxlsx::setRowHeights(
        wb,
        sheet   = sheet_name_1,
        rows    = 1,
        heights = 50
      )
      
      # 7. Column widths (sheet 1)
      openxlsx::setColWidths(
        wb,
        sheet  = sheet_name_1,
        cols   = 1:ncol(locality_data),
        widths = 30
      )
      
      # 8. Second sheet with ana_base_data filtered the same way
      sheet_name_2 <- "ANA Source Data"
      openxlsx::addWorksheet(wb, sheetName = sheet_name_2)
      
      if (nrow(base_locality_data) == 0) {
        openxlsx::writeData(
          wb,
          sheet = sheet_name_2,
          x = "NO DATA AVAILABLE FOR THIS LOCALITY",
          startRow = 1,
          startCol = 1
        )
      } else {
        openxlsx::writeData(
          wb,
          sheet = sheet_name_2,
          x = base_locality_data,
          startRow = 1,
          startCol = 1,
          headerStyle = NULL
        )
        
        openxlsx::addStyle(
          wb,
          sheet      = sheet_name_2,
          style      = header_style,
          rows       = 1,
          cols       = 1:ncol(base_locality_data),
          gridExpand = TRUE
        )
        
        openxlsx::setRowHeights(
          wb,
          sheet   = sheet_name_2,
          rows    = 1,
          heights = 50
        )
        
        openxlsx::setColWidths(
          wb,
          sheet  = sheet_name_2,
          cols   = 1:ncol(base_locality_data),
          widths = 30
        )
      }
      
      # 9. Save workbook
      openxlsx::saveWorkbook(wb, file, overwrite = TRUE)
    }
  )
  
  
  
  #################################### 4. LOCALITY REPORT DOCUMENT  ############
  
  # Food Section (Conclusion, Plausibility, and Triangulation)
  output$locality_report_header <- renderUI({
    req(input$locality_report)
    
    state <- ana_long %>%
      filter(`Administrative Level Label` == input$locality_report) %>%
      pull(State)
    state <- ifelse(length(state) == 0, "", unique(state)[1])
    
    # Use HTML() to allow for HTML formatting
    HTML(
      paste0(
        "<span style='font-size: 35px; font-weight: bold;'>", input$locality_report, "</span>", 
        "<span style='font-size: 32px; font-weight: normal; color: grey;'> – ", state, "</span>"
      )
    )
  })

  # Helper function to get the field value for the selected locality
  get_field <- function(loc, ind, system = NULL) {
    filter_conditions <- ana_long %>%
      filter(`Administrative Level Label` == loc, indicator == ind)
    
    # If a system is provided, further filter by system
    if (!is.null(system)) {
      filter_conditions <- filter_conditions %>%
        filter(System == system)
    }
    
    out <- filter_conditions %>%
      pull(value)
    if (length(out) == 0) return("—")
    out[1]
  }
  
  # Define conclusion colors
  conclusion_colors <- c(
    "C1: Strong Interaction" = "#FF9F9A",
    "C2: Limited or indirect Interaction" = "#FFE4A3",
    "C3: No interaction" = "#C9F2A3",
    "C1: Systemic Failure" = "#FF9F9A", 
    "C2: Pockets of severity" = "#FFE4A3",
    "C3: Mitigation/no acute needs" = "#C9F2A3",
    "Inconclusive" = "#D3D3D3"
  )
  
  # Locality Report Header (Locality + State)
output$locality_report_details <- renderUI({
  req(input$locality_report)
  
  # Get values for Conclusion, Severity, Main driver, and Secondary driver
  conclusion <- get_field(input$locality_report, "Conclusion")
  severity <- get_field(input$locality_report, "Severity")
  main_driver <- get_field(input$locality_report, "Main driver")
  secondary_driver <- get_field(input$locality_report, "Secondary driver")
  plausibility <- get_field(input$locality_report, "Plausibility judgement")
  triangulation <- get_field(input$locality_report, "Strenght of evidence triangulation")
  conclusion_long <- get_field(input$locality_report, "Conclusion Long")
  
  # Use the existing color palette (adjust accordingly)
  conclusion_color <- pal_deepdive_levels(conclusion)  # Assuming pal_deepdive is a predefined palette
  
  # Create the UI elements
  tagList(
    # Add margin at the top to space from the title
    div(style = "margin-top: 20px;"),
    
    div(
      style = paste("display: flex; align-items: center; font-size: 16px; font-weight: bold; color:", conclusion_color, "; gap: 15px;"),  # Flexbox layout with gap
      conclusion,
      span(
        style = paste("margin-left: 10px; margin-right: 10px;"),  # Add spacing before and after the separator
        "•"  # Separator (dot or dash)
      ),
      span(
        style = paste("font-size: 16px; color:", conclusion_color, ";"),
        paste(severity, "Severity")
      )
    ),
    
    
    # Main and Secondary driver in the same row
    div(
      style = "display: flex; gap: 20px; margin-top: 20px;",
      div(
        p("Main Driver of Needs: ", tags$strong(main_driver))  # Make the value bold
      ),
      div(
        p("Secondary Driver of Needs: ", tags$strong(secondary_driver))  # Make the value bold
      )
    ),
    
    div(
      style = "display: flex; gap: 20px;",
      div(
        p("Conclusion Plausibility: ", tags$strong(plausibility))  
      ),
      div(
        p("Evidence Triangulation: ", tags$strong(triangulation))  
      )
    ),
    
    # Conclusion Long (below the drivers)
    div(
      style = "margin-top: 30px; margin-bottom: 20px;",  # Adjust the space around the conclusion long
      h4(tags$strong("1. Humanitarian Needs Summary")),
      p(conclusion_long)
    )
  )
})

  
  # Create system output with colored dots and information for each system
output$system_report_details <- renderUI({
  req(input$locality_report)
  
  systems <- c("Food system", "Water system", "Living Conditions system", "Health/nutrition system", "Health outcomes")
  
  
  # Creating the UI elements for each system
  output_list <- lapply(systems, function(system_name) {
    # Get values for the system
    conclusion <- get_field(input$locality_report, "Conclusion", system_name)
    conclusion_summary <- get_field(input$locality_report, "Conclusion summary", system_name)
    plausibility <- get_field(input$locality_report, "Plausibility judgement", system_name)
    triangulation <- get_field(input$locality_report, "Strenght of evidence triangulation", system_name)
    
    # Determine the color for the conclusion
    color <- conclusion_colors[conclusion]
    
    # Create the UI elements for each system
    div(
      style = "margin-bottom: 30px;",
      
      # Title line: colored dot + system name
      div(
        style = "display: flex; align-items: center; margin-bottom: 10px;",
        div(
          style = paste("width: 12px; height: 12px; background-color:", color, "; border-radius: 50%; margin-right: 10px;")
        ),
        div(
          style = "font-weight: bold; flex-grow: 1;",
          system_name
        )
      ),
      
      # Conclusion summary (on a new line)
      div(
        style = "margin-bottom: 10px;",
        p(conclusion_summary)
      ),
      
      # Plausibility judgement and triangulation on the same line
      div(
        style = "display: flex; gap: 20px;",
        div(
          p("Plausibility: ", tags$strong(plausibility))  # Make the value bold
        ),
        div(
          p("Triangulation: ", tags$strong(triangulation))  # Make the value bold
        )
      ),
      
      # Separator between systems
      tags$hr(style = "margin-top: 10px;")
    )
  })
  
})

output$across_system_report_details <- renderUI({
  req(input$locality_report)
  
  # Define combinations of systems
  combinations <- c(
    "Food x Health outcomes", "Food x Water system", "Food x Living conditions", 
    "Food x Water system", "Health system x Health outcomes", 
    "Living conditions x Health outcomes", "Water x Health outcomes", 
    "Water x Living conditions"
  )
  
  
  # Creating the UI elements for each combination
  output_list <- lapply(combinations, function(combination_name) {
    # Get values for the combination
    conclusion <- get_field(input$locality_report, "Conclusion", combination_name)
    conclusion_summary <- get_field(input$locality_report, "Conclusion summary", combination_name)
    plausibility <- get_field(input$locality_report, "Plausibility judgement", combination_name)
    triangulation <- get_field(input$locality_report, "Strenght of evidence triangulation", combination_name)
    
    # Define the color for the conclusion
    color <- conclusion_colors[conclusion]
    
    # Create the UI elements for each combination
    div(
      style = "margin-bottom: 30px;",
      
      # Title line: colored dot + combination name
      div(
        style = "display: flex; align-items: center; margin-bottom: 10px;",
        div(
          style = paste("width: 12px; height: 12px; background-color:", color, "; border-radius: 50%; margin-right: 10px;")
        ),
        div(
          style = "font-weight: bold; flex-grow: 1;",
          combination_name
        )
      ),
      
      # Conclusion summary (on a new line)
      div(
        style = "margin-bottom: 10px;",
        p(conclusion_summary)
      ),
      
      # Plausibility judgement and triangulation on the same line
      div(
        style = "display: flex; gap: 20px;",
        div(
          p("Plausibility: ", tags$strong(plausibility))  # Make the value bold
        ),
        div(
          p("Triangulation: ", tags$strong(triangulation))  # Make the value bold
        )
      ),
      
      # Separator between combinations
      tags$hr(style = "margin-top: 10px;")
    )
  })
  
})

######################################## FULL DATA DOWNLOAD ###################
  
  # Download handler: copy the original file as-is to preserve formatting
output$download_ana_file <- downloadHandler(
  filename = function() {
    "99_ANA_Flagging_Deep_Analysis.xlsx"
  },
  content = function(file) {
    file.copy("input/99_ANA_Flagging_Deep_Analysis.xlsx", file, overwrite = TRUE)
  }
)

# Read first sheet once for preview
ana_first_sheet <- readxl::read_xlsx(
  "input/99_ANA_Flagging_Deep_Analysis.xlsx",
  sheet = 1
)

# DT preview of first sheet
output$ana_preview <- DT::renderDataTable({
  DT::datatable(
    ana_summary,
    options = list(
      pageLength = 7,
      scrollX   = TRUE
    ),
    class = "stripe hover compact"
  )
})



}
  


shinyApp(ui, server)
