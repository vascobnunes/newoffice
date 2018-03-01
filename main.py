import os
import subprocess
import json
import csv
import time
import simplekml
import imp
import sys
from xlrd import open_workbook


#set file where variables are defined
variablesfile= os.path.join(os.getcwd(),'variables.txt')


#function to read variables from txt file
def getVarFromFile(filename):
    f = open(filename)
    global data
    data = imp.load_source('', '', f)
    f.close()


#function to get coordinates from general address
def get_coordinates(address, api_key):
    message = "https://maps.googleapis.com/maps/api/geocode/json?address="+address+"&key="
    print("curl -v \"" + message + api_key + "\"")
    process = subprocess.Popen("curl -v \"" + message + api_key + "\"", stdout=subprocess.PIPE)
    out, err = process.communicate()
    d = json.loads(out)
    try:
        coordinates = [d['results'][0]['geometry']['location']['lng'], d['results'][0]['geometry']['location']['lat']]
    except:
        coordinates = 0
    return coordinates


#main function
def main():
    #get variables from txt file
    getVarFromFile(variablesfile)
    api_key = data.api_key
    export_csv_file = data.export_csv_file
    export_kml_file = data.export_kml_file
    addresses = data.addresses
    icon1 = data.icon1
    icon2 = data.icon2
    new_office_locations = data.new_office_locations
    person_type = data.person_type
    rush_hour = data.rush_hour
    if rush_hour == None or rush_hour == "": rush_hour = str(time.time())[0:10]

    kml=simplekml.Kml()

    #read addresses from excel file
    wb = open_workbook(addresses)
    employees_addresses = []
    for s in wb.sheets():
        for row in range(1, s.nrows):
            value = s.cell(row,1).value
            try: value = str(int(value))
            except: pass
            if value != "":
                employees_addresses.append(value)

    #write results to csv file
    with open(export_csv_file, 'w') as myfile:
        wr = csv.writer(myfile, quoting=csv.QUOTE_MINIMAL)
        wr.writerow(["office_location", "distance (m)", "duration driving (mins)", "duration driving rush hours (mins)", "duration transports (mins)"])
        check_employee = 0
        #cycle through interest locations
        for no in new_office_locations:
            distances = []
            durations_car = []
            durations_traffic = []
            durations_bus = []
            #cycle through people's addresses (retrieved from excel file)
            for ea in employees_addresses:
                #composing google maps api url
                message1 = "https://maps.googleapis.com/maps/api/distancematrix/json?origins="+ea+"&destinations="+no[0]+"&key="
                # message2 = "https://maps.googleapis.com/maps/api/distancematrix/json?origins="+ea+"&destinations="+no[0]+"&avoid=highways&mode=bicycling&key="
                message2 = "https://maps.googleapis.com/maps/api/distancematrix/json?origins="+ea+"&destinations="+no[0]+"&departure_time="+rush_hour+"&traffic_model=best_guess&key="
                message3 = "https://maps.googleapis.com/maps/api/distancematrix/json?origins=" + ea + "&destinations=" + no[0] + "&departure_time="+rush_hour+"&mode=transit&key="
                messages = [message1, message2, message3]
                results = []
                control = 0
                #cycle through different types of requests
                for msg in messages:
                    print("curl -v \""+msg+api_key+"\"")
                    process = subprocess.Popen("curl -v \""+msg+api_key+"\"", stdout=subprocess.PIPE)
                    out, err = process.communicate()
                    d = json.loads(out)
                    try:
                        distance = d['rows'][0]['elements'][0]['distance']['value']
                    except:
                        distance = 0
                    try:
                        duration = int(d['rows'][0]['elements'][0]['duration']['value'])
                    except:
                        duration = 0
                    if control == 0:
                        results.append(distance)
                    results.append(duration)
                    control += 1
                #register person's coordinates to kml
                coords = get_coordinates(ea, api_key)
                if check_employee == 0:
                    colaborator = kml.newpoint(name=person_type, coords=[(coords)])
                    colaborator.style.iconstyle.icon.href = icon2
                #append result to interest location
                if results[0] > 0: distances.append(results[0])
                if results[1] > 0: durations_car.append(results[1])
                if results[3] > 0: durations_bus.append(results[3])
                if results[2] > 0: durations_traffic.append(results[2])
                #write to csv
                wr.writerow([no[1], int(results[0]), int(results[1])/60, int(results[2])/60, results[3]/60])
            check_employee += 1
            #do some stats for interest location
            if len(distances) > 0:
                avg_distance = round(sum(distances) / len(distances)/1000, 0)
            else:
                avg_distance = "n/a"
            if len(durations_car) > 0:
                avg_durations_car = round(sum(durations_car) / len(durations_car) / 60, 0)
            else:
                avg_durations_car = "n/a"
            if len(durations_traffic) > 0:
                avg_durations_traffic = round(sum(durations_traffic) / len(durations_traffic) / 60, 0)
            else:
                avg_durations_traffic = "n/a"
            if len(durations_bus) > 0:
                avg_durations_bus = round(sum(durations_bus) / len(durations_bus) / 60, 0)
            else:
                avg_durations_bus = "n/a"
            #register interest location with stats to kml
            no_lat = float(no[0].split(",")[0])
            no_lng = float(no[0].split(",")[1])
            new_office = kml.newpoint(name=no[1], coords=[(no_lng, no_lat)])
            new_office.description = "Average distance: " + str(avg_distance)+"km; Average duration by car: " + str(avg_durations_car) +"min; Average duration by public transports: "+ str(avg_durations_bus)+"min"
            # new_office.description = "Average distance: " + str(avg_distance)+"km; Average duration by car: " + str(avg_durations_car) +"min; Average duration by car in rush hour: "+ str(avg_durations_traffic) + "min; Average duration by public transports: "+ str(avg_durations_bus)+"min"
            new_office.style.iconstyle.icon.href = icon1
    #save kml
    kml.save(export_kml_file)


if __name__ == "__main__":
    try:
        main()
    except:
        print("Unexpected error:", sys.exc_info()[1])