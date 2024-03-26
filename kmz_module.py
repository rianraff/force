import pandas as pd
import random
import simplekml
from fastkml import kml
from lxml import html
import pandas as pd
from zipfile import ZipFile
from pyproj import Transformer
from shapely.ops import transform
from shapely.geometry import Point
from pyproj import Geod
from shapely import LineString
from shapely.ops import substring
from shapely import MultiLineString
import warnings
import pandas as pd
import os
import re
from openpyxl import load_workbook

warnings.filterwarnings("ignore", category=RuntimeWarning)

def get_placemark(file_path):
  kmz = ZipFile(file_path, 'r')
  doc = kmz.open('doc.kml', 'r').read()
  kml_file = html.fromstring(doc)
  placemark_dict = {}
  for pm in kml_file.cssselect("Folder"):
    folder_name = pm.cssselect("name")[0].text_content()
    if folder_name not in placemark_dict:
      placemark_dict[folder_name] = pm.cssselect("Placemark")
    elif folder_name in placemark_dict:
      placemark_dict[folder_name + "_2"] = pm.cssselect("Placemark")
  return placemark_dict

# @title run this (just once)
def get_homepass_folder(placemark_dict):
  folder = []
  for hp_folder in ["HOME", "HOME-BIZ", "BIZ-HOME", "BIZ"]:
    if hp_folder in placemark_dict.keys():
      folder.append(hp_folder)
  return folder

def long_lat_mapping(coor):
  longitude = []
  latitude = []
  coordinates = coor
  for o, x in enumerate(coordinates):
    if o % 2 == 0:
      longitude.append(x)
    else:
      latitude.append(x)
  return longitude, latitude

def to_df(placemark_dict, parse_simple=False, mapping=False):
  data = []
  for pm in placemark_dict:
    for name in pm.cssselect("name"):
      name = name.text_content()
    for coord in pm.cssselect("coordinates"):
      coords = [float(i.replace("0 ", "")) for i in coord.text_content().strip().split(",") if len(i)>=1][:-1]
      if mapping:
        longitude, latitude = long_lat_mapping(coords)
        coords = list(zip(longitude, latitude))
    if parse_simple:
      for simple in pm.cssselect("SimpleData"):
        if simple.attrib['name'] == "FAT_CODE":
          FAT_CODE = simple.text_content()
        if simple.attrib["name"] == "HOMEPASS_ID":
          homepass_id = simple.text_content()
    if parse_simple:
      data.append([name, coords, FAT_CODE, homepass_id])
    else:
      data.append([name, coords])
  if parse_simple:
    df = pd.DataFrame(data=data, columns=["Name", "Coordinates", "FAT_CODE", "homepass_id"])
  else:
    df = pd.DataFrame(data=data, columns=["Name", "Coordinates"])
  return df

def make_circle_coord(longitude, latitude, radius=2):
  local_azimuthal_projection = "+proj=aeqd +R=6371000 +units=m +lat_0={} +lon_0={}".format(latitude, longitude)
  wgs84_to_aeqd = Transformer.from_proj('+proj=longlat +datum=WGS84 +no_defs', local_azimuthal_projection)
  aeqd_to_wgs84 = Transformer.from_proj(local_azimuthal_projection, '+proj=longlat +datum=WGS84 +no_defs')
  # Get polygon with lat lon coordinates
  point_transformed = Point(wgs84_to_aeqd.transform(longitude, latitude))

  buffer = point_transformed.buffer(radius)
  circle = transform(aeqd_to_wgs84.transform, buffer)
  return circle

def mapping_hp_to_pole(hp_df, pole_df):
  hp_df = hp_df
  geod = Geod(ellps="WGS84")
  pole_list = []
  for _, hp in hp_df.iterrows():
    hp_point = Point(hp["Coordinates"][0], hp["Coordinates"][1])
    distance_list = []
    for pole_index, pole in pole_df.iterrows():
      pole_point = Point(pole["Coordinates"][0], pole["Coordinates"][1])
      line = LineString([hp_point, pole_point])
      distance_list.append(geod.geometry_length(line))
    lowest_distance = min(distance_list)
    index_of_lowest_value = distance_list.index(lowest_distance)
    pole_list.append(index_of_lowest_value)
  hp_df["Pole_index"] = pole_list
  return hp_df

def check_pole_to_hp(placemark_dict, pole_df):
  hp_folder_name = get_homepass_folder(placemark_dict)
  pole_to_hp_35m = []
  pole_to_hp_35m_coords = []
  for name in hp_folder_name:
    homepass_df = to_df(placemark_dict[name], parse_simple=True)
    geod = Geod(ellps="WGS84")
    for _, hp in homepass_df.iterrows():
      hp_point = Point(hp["Coordinates"][0], hp["Coordinates"][1])
      distance_list = []
      for _, pole in pole_df.iterrows():
        pole_point = Point(pole["Coordinates"][0], pole["Coordinates"][1])
        line = LineString([hp_point, pole_point])
        distance_list.append(geod.geometry_length(line))
      lowest_distance = min(distance_list)
      if lowest_distance > 35:
        pole_to_hp_35m.append("{} {}".format(name, hp["Name"]))
        pole_to_hp_35m_coords.append(str((hp["Coordinates"][0], hp["Coordinates"][1])))   
    return pole_to_hp_35m, pole_to_hp_35m_coords
  
def getAllHP(file_path):
    placemark_dict = get_placemark(file_path)  
    hp_folder_name = get_homepass_folder(placemark_dict)
    all_homepass_df = pd.DataFrame()  # DataFrame kosong untuk menyimpan semua data

    for name in hp_folder_name:
        homepass_df = to_df(placemark_dict[name], parse_simple=True)
        all_homepass_df = pd.concat([all_homepass_df, homepass_df], ignore_index=True)  # Gabungkan DataFrame ke DataFrame besar

    return all_homepass_df

def getAllFAT(file_path):
    placemark_dict = get_placemark(file_path)
    fat_df = to_df(placemark_dict["FAT"], parse_simple=False)    

    return fat_df

def check_fat_to_hp(placemark_dict, pole_df, fat_df, cable_df):
  fat_to_hp_150m = []
  fat_to_hp_coords = []
  hp_folder_name = get_homepass_folder(placemark_dict)
  has_sling = False
  try:
    sling_df = to_df(placemark_dict["SLINGWIRE"], parse_simple=False, mapping=True)
    has_sling = True
  except:
    has_sling = False
  for name in hp_folder_name:
    homepass_df = to_df(placemark_dict[name], parse_simple=True)
    geod = Geod(ellps="WGS84")
    hp_df = mapping_hp_to_pole(homepass_df, pole_df)
    count_case1 = 0
    count_case2 = 0
    count_case3 = 0
    case3_name = []
    case3_coord = []
    count_case4 = 0
    case4_name = []
    case4_coord = []
    for _, hp in hp_df.iterrows():
      fat = fat_df[fat_df["Name"] == hp["FAT_CODE"]]
      if len(fat) == 0:
        fat_to_hp_150m.append("there is no fat {} in kmz file".format(hp["FAT_CODE"]))
        fat_to_hp_coords.append(hp["Coordinates"])
        continue
      pole_coordinates = pole_df["Coordinates"].iloc[hp["Pole_index"]]
      hp_point = Point(hp["Coordinates"][0], hp["Coordinates"][1])
      fat_point = Point(fat["Coordinates"].item()[0], fat["Coordinates"].item()[1])
      pole_point = Point(pole_coordinates[0], pole_coordinates[1])
      pole_area = make_circle_coord(pole_coordinates[0], pole_coordinates[1], radius=3)
      dropwire_line = LineString([hp_point, pole_point])
      dropwire_length = geod.geometry_length(dropwire_line)
      has_sling_line = False
      if has_sling:
        for _, sling in sling_df.iterrows():
          if len(sling["Coordinates"]) > 1:
            tmp_sling_line = LineString(sling["Coordinates"])
            if tmp_sling_line.intersects(pole_area):
              has_sling_line = tmp_sling_line.intersects(pole_area)
      tmp_distance = []
      for _, cable in cable_df.iterrows():
        tmp_cable_line = LineString(cable["Coordinates"])
        dist_in_cable = tmp_cable_line.line_locate_point(pole_point)
        point_in_cable = tmp_cable_line.line_interpolate_point(dist_in_cable)
        tmp_line = LineString([point_in_cable, pole_point])
        tmp_distance.append(geod.geometry_length(tmp_line))
      lowest_distance = min(tmp_distance)
      index_of_lowest_value = tmp_distance.index(lowest_distance)
      cable = cable_df.iloc[index_of_lowest_value]
      cable_line = LineString(cable["Coordinates"])
      pole_to_fat_line = substring(cable_line, start_dist=cable_line.line_locate_point(fat_point), end_dist=cable_line.line_locate_point(pole_point))
      pole_to_fat_length = geod.geometry_length(pole_to_fat_line)
      distance = dropwire_length + pole_to_fat_length
      if has_sling_line:
        pole_to_fat_line_air = LineString([pole_point, fat_point])
        pole_to_fat_air_length = geod.geometry_length(pole_to_fat_line_air)
        distance = dropwire_length + pole_to_fat_air_length
      if distance > 150:
        count_case2 += 1
        pole_to_fat_line_air = LineString([pole_point, fat_point])
        pole_to_fat_air_length = geod.geometry_length(pole_to_fat_line_air)
        distance_case2 = dropwire_length + pole_to_fat_air_length
        if distance_case2 <= 150:
          count_case3 += 1
          fat_to_hp_150m.append("{} {} is > 150m and but in airhead < 150m".format(name, hp["Name"]))
          fat_to_hp_coords.append(str((hp["Coordinates"][0], hp["Coordinates"][1])))
        elif distance_case2 > 150:
          count_case4 += 1
          fat_to_hp_150m.append("{} {} is > 150m".format(name, hp["Name"]))
          fat_to_hp_coords.append(str((hp["Coordinates"][0], hp["Coordinates"][1])))
  return fat_to_hp_150m, fat_to_hp_coords

def is_fat_contain_pole(df_pole, df_fat):
  fat_pole = {}
  does_not_contain_pole = []
  for _, row_fat in df_fat.iterrows():
    circle = make_circle_coord(row_fat["Coordinates"][0], row_fat['Coordinates'][1], radius = 2)
    for _, row_pole in df_pole.iterrows():
      p = Point(row_pole["Coordinates"][0], row_pole["Coordinates"][1])
      condition = circle.contains(p)
      if condition:
        if row_fat["Name"] not in fat_pole.keys():
          fat_pole[row_fat["Name"]] = [row_pole["Name"]]
        else:
          fat_pole[row_fat["Name"]].append(row_pole["Name"])
      else:
        if row_fat["Name"] not in fat_pole.keys():
          fat_pole[row_fat["Name"]] = []
        else:
          continue
  for key, values in fat_pole.items():
    if len(values) == 1:
      continue
    else:
      does_not_contain_pole.append(key)
  return does_not_contain_pole

def is_fdt_contain_pole(df_pole, df_fdt):
  fdt_pole = {}
  does_not_contain_pole = []
  for _, row_fdt in df_fdt.iterrows():
    circle = make_circle_coord(row_fdt["Coordinates"][0], row_fdt['Coordinates'][1], radius = 2)
    for _, row_pole in df_pole.iterrows():
      p = Point(row_pole["Coordinates"][0], row_pole["Coordinates"][1])
      condition = circle.contains(p)
      if condition:
        if row_fdt["Name"] not in fdt_pole.keys():
          fdt_pole[row_fdt["Name"]] = [row_pole["Name"]]
        else:
          fdt_pole[row_fdt["Name"]].append(row_pole["Name"])
      else:
        if row_fdt["Name"] not in fdt_pole.keys():
          fdt_pole[row_fdt["Name"]] = []
        else:
          continue
  for key, values in fdt_pole.items():
    if len(values) == 1:
      continue
    else:
      does_not_contain_pole.append(key)
  return does_not_contain_pole

def check_duplicate_hpid(homepass_df):
  mask = homepass_df["homepass_id"].duplicated()
  duplicate = homepass_df[mask]
  for i, j in duplicate.iterrows():
    print("house number {} with homepass id {} has duplicate".format(j["Name"], j["homepass_id"]))

def check_distribution_cable_connect_to_pole(df_cable, df_pole):
  no_intersection_count = {}
  for _, pole in df_pole.iterrows():
    pole_area = make_circle_coord(pole["Coordinates"][0], pole["Coordinates"][1], radius=2)
    for _, cable in df_cable.iterrows():
      cable_line = LineString(cable['Coordinates'])
      intersect = cable_line.intersects(pole_area)
      if not intersect:
        if pole["Name"] not in no_intersection_count:
          no_intersection_count[pole["Name"]] = 1
        else:
          no_intersection_count[pole["Name"]] += 1
  cable_not_in_distribution = []
  for name, count in no_intersection_count.items():
    if count == len(df_cable):
      cable_not_in_distribution.append(name)
  filtered_df = df_pole[df_pole['Name'].isin(cable_not_in_distribution)]
  return filtered_df

def check_pole_has_sling(df_cable, df_pole, df_sling):
  pole_not_in_distribution_df = check_distribution_cable_connect_to_pole(df_cable, df_pole)
  no_intersection_count = {}
  for _, pole in pole_not_in_distribution_df.iterrows():
    pole_area = make_circle_coord(pole["Coordinates"][0], pole["Coordinates"][1], radius=2)
    for _, sling in df_sling.iterrows():
      if len(sling["Coordinates"]) > 1:
        sling_line = LineString(sling["Coordinates"])
        intersect = sling_line.intersects(pole_area)
      else:
        continue
      if not intersect:
        if pole["Name"] not in no_intersection_count:
          no_intersection_count[pole["Name"]] = 1
        else:
          no_intersection_count[pole["Name"]] += 1
  pole_not_in_sling = []
  for name, count in no_intersection_count.items():
    if count == len(df_sling):
      pole_not_in_sling.append(name)
  filtered_df = pole_not_in_distribution_df[pole_not_in_distribution_df['Name'].isin(pole_not_in_sling)]
  pole_has_no_cable = []
  pole_has_no_cable_coords = []
  for index, row in filtered_df.iterrows():
    pole_has_no_cable.append(row["Name"])
    pole_has_no_cable_coords.append(str(row["Coordinates"]))
  return pole_has_no_cable, pole_has_no_cable_coords

def check_row_has_value(df, column_name, desired_value):
    for index, row in df.iterrows():
        if row[column_name] != desired_value:
            return False
    return True

def kmzCheck(file_path, cluster, checking_date, checking_time):   
    hpdb_col = [
    'ACQUISITION_CLASS',
    'ACQUISITION_TIER',
    'BUILDING_TYPE',
    'OWNERSHIP',
    'VENDOR_NAME',
    'ZIP_CODE',
    'REGION',
    'CITY',
    'CITY_CODE',
    'DISTRICT',
    'SUB_DISTRICT',
    'FAT_CODE',
    'FAT_LONGITUDE',
    'FAT_LATITUDE',
    'BUILDING_LATITUDE',
    'BUILDING_LONGITUDE',
    'HOMEPASS_ID',
    'MOBILE_REGION',
    'MOBILE_CLUSTER',
    'CITY_GROUP']
    kmz_col = ['Pole to FAT', 'Pole to FDT', 'HP to pole 35m', 'Coordinate HP to pole 35m', 'HP to FAT 150m', 'Coordinate HP to FAT 150m',
              "Pole not in Distribution and Sling", "Coordinate Pole not in Distribution and Sling"]
    log_col = ['Cluster ID', 'Checking Date', 'Checking Time', "Status"]

    placemark_dict = get_placemark(file_path)

    has_sling = False
    has_pole = False
    has_fat = False
    has_fdt = False
    has_cable = False

    try:
      pole_df = to_df(placemark_dict["POLE"], parse_simple=False)
      has_pole = True
    except:
      print("no folder POLE")
    try:
      fat_df = to_df(placemark_dict["FAT"], parse_simple=False)
      has_fat = True
    except:
      print("no folder FAT")
    try:
      fdt_df = to_df(placemark_dict["FDT"], parse_simple=False)
      has_fdt = True
    except:
      print("no folder FDT")
    try:
      cable_df = to_df(placemark_dict["CABLE DISTRIBUTION"], parse_simple=False, mapping=True)
      has_cable = True
    except:
      print("no folder CABLE DISTRIBUTION")
    try:
      sling_df = to_df(placemark_dict["SLINGWIRE"], parse_simple=False, mapping=True)
      has_sling = True
    except:
      print("no folder SLINGWIRE")

    if has_pole:
      pole_to_hp, pole_to_hp_coords = check_pole_to_hp(placemark_dict, pole_df)
    else:
      pole_to_hp = ["No POLE folder in kmz"]
      pole_to_hp_coords = ["-"]

    if has_pole and has_fat and has_cable: 
      fat_to_hp, fat_to_hp_coords = check_fat_to_hp(placemark_dict, pole_df, fat_df, cable_df)   
    else:
      fat_to_hp = ["No POLE or FAT or CABLE DISTRIBUTION folder in kmz"]
      fat_to_hp_coords = ["-"]

    if has_pole and has_fat: 
      fat_to_pole = is_fat_contain_pole(pole_df, fat_df)
    else:
      fat_to_pole = ["No POLE or FAT folder in kmz"]

    if has_pole and has_fdt:
      fdt_to_pole = is_fdt_contain_pole(pole_df, fdt_df)
    else:
      fdt_to_pole = ["No FDT folder in kmz"]

    if has_sling:
      if has_cable and has_pole:
        pole_has_no_cable, pole_has_no_cable_coords = check_pole_has_sling(cable_df, pole_df, sling_df)
      else:
        pole_has_no_cable = ["No POLE or CABLE DISTRIBUTION folder in kmz"]
        pole_has_no_cable_coords = ["-"]
    else:
      pole_has_no_cable = []
      pole_has_no_cable_coords = []


    # Create a DataFrame with column names from column_names
    kmz_df = pd.DataFrame(columns=log_col + kmz_col + hpdb_col)
  
    if len(fat_to_pole) == 0:
      row_temp = {}
      row_temp["Cluster ID"] = cluster
      row_temp["Checking Date"] = checking_date
      row_temp["Checking Time"] = checking_time
      row_temp["Status"] = "OK"
      for col_name in kmz_col:
        if col_name == "Pole to FAT":
          row_temp[col_name] = "OK"
        else:
          row_temp[col_name] = "-"
      for col_name in hpdb_col:
        row_temp[col_name] = "-"
      new_row_df = pd.DataFrame([row_temp])
      kmz_df = kmz_df._append(new_row_df, ignore_index=True)
    else:
      for i in fat_to_pole:
        row_temp = {}
        row_temp["Cluster ID"] = cluster
        row_temp["Checking Date"] = checking_date
        row_temp["Checking Time"] = checking_time
        row_temp["Status"] = "REVISE"
        for col_name in kmz_col:
          if col_name == "Pole to FAT":
            row_temp[col_name] = i
          else:
            row_temp[col_name] = "-"
        for col_name in hpdb_col:
          row_temp[col_name] = "-"
        new_row_df = pd.DataFrame([row_temp])
        kmz_df = kmz_df._append(new_row_df, ignore_index=True)
        
    
    if len(fdt_to_pole) == 0:
      row_temp = {}
      row_temp["Cluster ID"] = cluster
      row_temp["Checking Date"] = checking_date
      row_temp["Checking Time"] = checking_time
      row_temp["Status"] = "OK"
      for col_name in kmz_col:
        if col_name == "Pole to FDT":
          row_temp[col_name] = "OK"
        else:
          row_temp[col_name] = "-"
      for col_name in hpdb_col:
        row_temp[col_name] = "-"
      new_row_df = pd.DataFrame([row_temp])
      kmz_df = kmz_df._append(new_row_df, ignore_index=True)
    else:
      for i in fdt_to_pole:
        row_temp = {}
        row_temp["Cluster ID"] = cluster
        row_temp["Checking Date"] = checking_date
        row_temp["Checking Time"] = checking_time
        row_temp["Status"] = "REVISE"
        for col_name in kmz_col:
          if col_name == "Pole to FDT":
            row_temp[col_name] = i
          else:
            row_temp[col_name] = "-"
        for col_name in hpdb_col:
          row_temp[col_name] = "-"
        new_row_df = pd.DataFrame([row_temp])
        kmz_df = kmz_df._append(new_row_df, ignore_index=True)

    if len(pole_to_hp) == 0:
      row_temp = {}
      row_temp["Cluster ID"] = cluster
      row_temp["Checking Date"] = checking_date
      row_temp["Checking Time"] = checking_time
      row_temp["Status"] = "OK"
      for col_name in kmz_col:
        if col_name == "HP to pole 35m" or col_name == 'Coordinate HP to pole 35m':
          row_temp[col_name] = "OK"
        else:
          row_temp[col_name] = "-"
      for col_name in hpdb_col:
        row_temp[col_name] = "-"
      new_row_df = pd.DataFrame([row_temp])
      kmz_df = kmz_df._append(new_row_df, ignore_index=True)
    else:
      for i, j in zip(pole_to_hp, pole_to_hp_coords):
        row_temp = {}
        row_temp["Cluster ID"] = cluster
        row_temp["Checking Date"] = checking_date
        row_temp["Checking Time"] = checking_time
        row_temp["Status"] = "REVISE"
        for col_name in kmz_col:
          if col_name == "HP to pole 35m":
            row_temp[col_name] = i
          elif col_name == 'Coordinate HP to pole 35m':
            row_temp[col_name] = j
          else:
            row_temp[col_name] = "-"
        for col_name in hpdb_col:
          row_temp[col_name] = "-"
        new_row_df = pd.DataFrame([row_temp])
        kmz_df = kmz_df._append(new_row_df, ignore_index=True)

    if len(fat_to_hp) == 0:
      row_temp = {}
      row_temp["Cluster ID"] = cluster
      row_temp["Checking Date"] = checking_date
      row_temp["Checking Time"] = checking_time
      row_temp["Status"] = "OK"
      for col_name in kmz_col:
        if col_name == "HP to FAT 150m" or col_name == 'Coordinate HP to FAT 150m':
          row_temp[col_name] = "OK"
        else:
          row_temp[col_name] = "-"
      for col_name in hpdb_col:
        row_temp[col_name] = "-"
      new_row_df = pd.DataFrame([row_temp])
      kmz_df = kmz_df._append(new_row_df, ignore_index=True)
    else:
      for i, j in zip(fat_to_hp, fat_to_hp_coords):
        row_temp = {}
        row_temp["Cluster ID"] = cluster
        row_temp["Checking Date"] = checking_date
        row_temp["Checking Time"] = checking_time
        row_temp["Status"] = "REVISE"
        for col_name in kmz_col:
          if col_name == "HP to FAT 150m":
            row_temp[col_name] = i
          elif col_name == 'Coordinate HP to FAT 150m':
            row_temp[col_name] = j
          else:
            row_temp[col_name] = "-"
        for col_name in hpdb_col:
          row_temp[col_name] = "-"
        new_row_df = pd.DataFrame([row_temp])
        kmz_df = kmz_df._append(new_row_df, ignore_index=True)

    if len(pole_has_no_cable) == 0:
      row_temp = {}
      row_temp["Cluster ID"] = cluster
      row_temp["Checking Date"] = checking_date
      row_temp["Checking Time"] = checking_time
      row_temp["Status"] = "OK"
      for col_name in kmz_col:
        if col_name == "Pole not in Distribution and Sling" or col_name == 'Coordinate Pole not in Distribution and Sling':
          row_temp[col_name] = "OK"
        else:
          row_temp[col_name] = "-"
      for col_name in hpdb_col:
        row_temp[col_name] = "-"
      new_row_df = pd.DataFrame([row_temp])
      kmz_df = kmz_df._append(new_row_df, ignore_index=True)
    else:
      for i, j in zip(pole_has_no_cable, pole_has_no_cable_coords):
        row_temp = {}
        row_temp["Cluster ID"] = cluster
        row_temp["Checking Date"] = checking_date
        row_temp["Checking Time"] = checking_time
        row_temp["Status"] = "REVISE"
        for col_name in kmz_col:
          if col_name == "Pole not in Distribution and Sling":
            row_temp[col_name] = i
          elif col_name == 'Coordinate Pole not in Distribution and Sling':
            row_temp[col_name] = j
          else:
            row_temp[col_name] = "-"
        for col_name in hpdb_col:
          row_temp[col_name] = "-"
        new_row_df = pd.DataFrame([row_temp])
        kmz_df = kmz_df._append(new_row_df, ignore_index=True)

    condition = check_row_has_value(kmz_df, "Status", "REVISE")

    summary_file_path = "Summary\Checking Summary.xlsx"
    if not os.path.exists(summary_file_path):
        kmz_df.to_excel(summary_file_path, index=False, engine='openpyxl')
    
    else:
        with pd.ExcelWriter(summary_file_path, 'openpyxl', mode='a',  if_sheet_exists="overlay") as writer:
            # fix line
            reader = pd.read_excel(summary_file_path, sheet_name="Sheet1")
            kmz_df.to_excel(writer, sheet_name="Sheet1", index=False, engine='openpyxl', header=False, startrow=len(reader)+1)

    print("Done")

    return condition