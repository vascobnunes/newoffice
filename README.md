# newoffice

Calculate average road distances and transit times from people's addresses to interest locations

Google api_keys required:
 - Distance Matrix (https://developers.google.com/maps/documentation/distance-matrix/)
 - Geocoding (https://developers.google.com/maps/documentation/geocoding/intro)
 
Request both keys for same project and insert last key in variables.txt

To execute, first set the variables.txt file:

`api_key = "AfzaSyD4dm5ikdfgQofetyGh1zqRI2JYM9ob4Xw"  #INSERT GOOGLE MAPS API KEY HERE`

`export_csv_file = "C:\\Users\\Public\\statsdata.csv"  #PATH TO CSV FILE`

`export_kml_file = "C:\\Users\\Public\\geodata.kml"  #PATH TO KML FILE`

`addresses = "C:\\Users\\Public\\persons.xlsx"  #PATH TO EXCEL FILE WITH ADDRESSES`

`icon1 = "http://famoesclubeatletico.com/Content/images/demo_logo.png"  #ICON FOR INTEREST LOCATIONS`

`icon2 = "http://cmauriflama.sp.gov.br/app/webroot/img/avatar-user.png"  #ICON FOR PEOPEL'S ADDRESSES`

`new_office_locations = [["38.784577,-9.213758", "Pavilhão Susana Barroso"], ["38.7945988,-9.1954576","Escola Moinhos da Arroja"]]  #INTEREST LOCATIONS LIST - [["lat,lon", "Title"], [...]]`

`person_type = "Colaborator"  #Tag for persons' locations`

`rush_hour = "1519913534"  #INSERT RUSH HOUR IN UNIX TIMESTAMP - ONLY FOR PREMIUM API KEYS`



