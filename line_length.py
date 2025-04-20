import geopandas as gpd
from shapely.geometry import Point, LineString, MultiLineString
import pandas as pd

# --- Load KML files ---
lines_gdf = gpd.read_file("Electric_Power_Transmission_Lines.kml", driver='KML')
points_gdf = gpd.read_file("Electric_Substations.kml", driver='KML')

# --- Reproject to metric CRS (UTM zone can be changed if needed) ---
lines_gdf = lines_gdf.to_crs(epsg=32633)
points_gdf = points_gdf.to_crs(epsg=32633)

# --- Helper to find nearest substation by geometry ---
def find_nearest(point, gdf_points):
    distances = gdf_points.geometry.distance(point)
    nearest_idx = distances.idxmin()
    return gdf_points.loc[nearest_idx]["Name"]

# --- Process each line (handle MultiLineString and LineString) ---
from_names = []
to_names = []
distances = []

for _, row in lines_gdf.iterrows():
    geom = row.geometry

    if isinstance(geom, MultiLineString):
        total_length = 0
        first_segment = list(geom.geoms)[0]
        last_segment = list(geom.geoms)[-1]
        start_point = Point(first_segment.coords[0])
        end_point = Point(last_segment.coords[-1])
        for segment in geom.geoms:
            total_length += segment.length
    elif isinstance(geom, LineString):
        start_point = Point(geom.coords[0])
        end_point = Point(geom.coords[-1])
        total_length = geom.length
    else:
        continue  # Skip non-line geometries

    from_sub = find_nearest(start_point, points_gdf)
    to_sub = find_nearest(end_point, points_gdf)

    from_names.append(from_sub)
    to_names.append(to_sub)
    distances.append(total_length)

# --- Create DataFrame and save to Excel ---
result_df = pd.DataFrame({
    "from_substation": from_names,
    "to_substation": to_names,
    "distance_m": distances
})

result_df.to_excel("substation_distances.xlsx", index=False)
