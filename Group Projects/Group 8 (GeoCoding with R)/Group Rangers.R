#Adding Geocode
library(tidyverse)
library(tidygeocoder)
library(sf)
library(mapview)
library(dplyr)
library(ggplot2)
library(leaflet)

gig <- read.csv("C:/Users/Open User.DESKTOP-R82ORKS/Desktop/Group Rangers/GIG Logistic.csv")
gig
names(gig)

sys <-  Sys.setenv(GOOGLEGEOCODE_API_KEY = 'AIzaSyAueq26b_j0W_S0bF0phOndVLTwL3nHy6k')
geo_code_tbl <- gig %>%
  
  #Geocode Address to lat/Lon
  tidygeocoder::geocode(
    address = Address,
    method = "google"
  )
geo_code_tbl
glimpse(geo_code_tbl)


#Plot Map Observation
gig2 <- geo_code_tbl %>%
  drop_na(long, lat) %>%
  st_as_sf(
    coords = c("long", "lat"),
    crs = 4326
  )
mapview(gig2)@map


# Insight 1: Regional Density of Offices
# Count offices by state
state_counts <- gig2 %>%
  st_drop_geometry() %>%
  count(State = State) 
state_counts


# Bar chart of office density
ggplot(state_counts, aes(x = reorder(State, n), y = n)) +
  geom_col(fill = "steelblue") +
  coord_flip() +
  labs(title = "Number of Offices by State", x = "State", y = "Office Count")


# Insight 2: Labelled location Per State 
leaflet(gig2) %>%
  addTiles() %>%
  addMarkers(
    popup = ~paste("Office:",Branches, "<br>State:", State)
  )


# Insight 3: Urban Clusters vs. Rural Spread
leaflet(gig2) %>%
  addTiles() %>%
  addMarkers(clusterOptions = markerClusterOptions())

