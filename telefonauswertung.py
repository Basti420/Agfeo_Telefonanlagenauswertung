import csv
import xlsxwriter
import glob
import os
import time
import sys
import datetime
import collections

GUELTIGE_EINTRAEGE = set()
ROW_COUNTER=0
DUPLICATE_COUNTER = 0
DURCHWAHL_NAME = {}
DURCHWAHL_ABGEHEND_ANZAHL = {}
DURCHWAHL_ABGEHEND_DAUER = {}
DURCHWAHL_DUPLIKATE_ANZAHL = {}
DURCHWAHL_ANKOMMEND_ANZAHL = {}
DURCHWAHL_ANKOMMEND_DAUER = {}
DURCHWAHL_DUPLIKATE_ZEIT = {}
NUMMERN = []
NICHT_AUSZUWERTENDE_DURCHWAHLEN = [0,14,19,20,21,26,28,30,37,43,50,51,60] # Durchwahlen die in der xlsx Datei ausgewertet werden sollen
FILE_NAME = "export.csv"
if not os.path.isfile(FILE_NAME):
	print("Exportdatei wurde nicht gefunden")
	time.sleep(5)
	sys.exit()
StartDatum = input("Startdatum (DD.MM.YYYY): " )
EndDatum = input("Enddatum (DD.MM.YYYY): ")
DATUM_START = datetime.datetime.strptime(StartDatum, '%d.%m.%Y')
DATUM_ENDE = datetime.datetime.strptime(EndDatum, '%d.%m.%Y')
with open(FILE_NAME, newline='') as csvfile:
	reader = csv.reader(csvfile, delimiter=',', quotechar='"')
	writer=csv.writer(open('export_duplikate.csv', 'w'), delimiter=',', quotechar='"',quoting=csv.QUOTE_ALL)

	for row in reader:
		datum = datetime.datetime.strptime(str(row[0]), '%Y-%m-%d %H:%M:%S')
		nummer = str(row[1])
		durchwahl = int(row[4])
		name = str(row[5])
		dauer_raw = row[9].split(':')
		dauer = datetime.timedelta(seconds=(int(dauer_raw[0])*60*60 + int(dauer_raw[1])*60 + int(dauer_raw[2])))
		richtung = int(row[12])
		if durchwahl not in DURCHWAHL_DUPLIKATE_ANZAHL:
			DURCHWAHL_DUPLIKATE_ANZAHL[durchwahl] = 0

		if durchwahl not in DURCHWAHL_DUPLIKATE_ZEIT:
			DURCHWAHL_DUPLIKATE_ZEIT[durchwahl]	= datetime.timedelta(seconds=0)

		key = (row[0], row[1], row[2], row[3], row[4], row[5], row[6], row[7], row[8], row[9])
		if key not in GUELTIGE_EINTRAEGE:
			GUELTIGE_EINTRAEGE.add(key)
			if durchwahl not in NICHT_AUSZUWERTENDE_DURCHWAHLEN:
				if DATUM_START.date() <= datum.date() and DATUM_ENDE.date() >= datum.date():

					if durchwahl not in DURCHWAHL_NAME:
						DURCHWAHL_NAME[durchwahl] = name
					if durchwahl not in DURCHWAHL_ABGEHEND_ANZAHL:
						DURCHWAHL_ABGEHEND_ANZAHL[durchwahl] = 0
					if durchwahl not in DURCHWAHL_ANKOMMEND_ANZAHL:
						DURCHWAHL_ANKOMMEND_ANZAHL[durchwahl] = 0
					if durchwahl not in DURCHWAHL_ABGEHEND_DAUER:
						DURCHWAHL_ABGEHEND_DAUER[durchwahl] = datetime.timedelta(seconds=0)
					if durchwahl not in DURCHWAHL_ANKOMMEND_DAUER:
						DURCHWAHL_ANKOMMEND_DAUER[durchwahl] = datetime.timedelta(seconds=0)
					if nummer != "":
						if richtung == 0:
							DURCHWAHL_ABGEHEND_ANZAHL[durchwahl] += 1
							DURCHWAHL_ABGEHEND_DAUER[durchwahl] += dauer
						else:
							DURCHWAHL_ANKOMMEND_ANZAHL[durchwahl] += 1
							DURCHWAHL_ANKOMMEND_DAUER[durchwahl] += dauer
					if nummer not in NUMMERN:
						NUMMERN.append(nummer)

		else:
			if durchwahl not in NICHT_AUSZUWERTENDE_DURCHWAHLEN:
				DUPLICATE_COUNTER += 1
				DURCHWAHL_DUPLIKATE_ANZAHL[durchwahl] += 1
				DURCHWAHL_DUPLIKATE_ZEIT[durchwahl] += dauer
				writer.writerow(row)
		ROW_COUNTER += 1
DURCHWAHL_NAME_ORDERED = collections.OrderedDict(sorted(DURCHWAHL_NAME.items()))
		
print("Von " + str(ROW_COUNTER) + " Einträgen, die Ausgewertet worden sind, wurden " + str(DUPLICATE_COUNTER) + " als Dupliakte erkannt.")
print("Duplikate wurden in die Datei export_duplikate.csv exportiert.")
print("")
print("Folgende Anschlüsse wurden ausgewertet:")

workbook = xlsxwriter.Workbook("Auswertung_"+DATUM_START.strftime('%d.%m.%Y')+"_"+DATUM_ENDE.strftime('%d.%m.%Y')+".xlsx")
worksheet = workbook.add_worksheet()

# Formatierung
fett_mittig = workbook.add_format({'bold': True})
fett_mittig.set_align('center')
zahlen = workbook.add_format()
zahlen.set_num_format('0.00')
zahlen.set_align('center')
zahlen_rot = workbook.add_format()
zahlen_rot.set_num_format('0.00')
zahlen_rot.set_align('center')
zahlen_rot.set_font_color('red')
zahlen_gruen = workbook.add_format()
zahlen_gruen.set_num_format('0.00')
zahlen_gruen.set_align('center')
zahlen_gruen.set_font_color('green')
mittig = workbook.add_format()
mittig.set_align('center')
zahlen_neu = workbook.add_format()
zahlen_neu.set_align('center')
worksheet.set_column('A:A', 20)
worksheet.set_column('B:T', 15)

# Inhalt
worksheet.write('A1', 'Telefonauswertung', fett_mittig)
worksheet.write('B1', 'von', fett_mittig)
worksheet.write('C1', str(DATUM_START.strftime('%d.%m.%Y')), fett_mittig)
worksheet.write('D1', 'bis', fett_mittig)
worksheet.write('E1', str(DATUM_ENDE.strftime('%d.%m.%Y')), fett_mittig)
worksheet.write('G1', "Tage:", fett_mittig)
worksheet.write('A2', 'Name', fett_mittig)
worksheet.write('A3', 'Durchwahl', fett_mittig)
worksheet.write('A5', 'Anzahl abgehend', fett_mittig)
worksheet.write('A6', 'Dauer abgehend', fett_mittig)
worksheet.write('A7', 'Dauer/Telefonat', fett_mittig)
worksheet.write('A9', 'Anzahl eingehend', fett_mittig)
worksheet.write('A10', 'Dauer eingehend', fett_mittig)
worksheet.write('A11', 'Dauer/Telefonat', fett_mittig)
worksheet.write('A13', 'Gesamtanzahl', fett_mittig)
worksheet.write('A14', 'Gesamtdauer', fett_mittig)
worksheet.write('A15', 'Dauer/Woche', fett_mittig)
 
ROW_COUNTER_XLSX=1
TAGE = ((DATUM_ENDE.date()-DATUM_START.date()).days+1)
worksheet.write('H1', str(TAGE), fett_mittig)
for key in DURCHWAHL_NAME_ORDERED:
		print(str(DURCHWAHL_NAME_ORDERED[key]) + " (DW: " + str(key) + ", Duplikate: " + str(DURCHWAHL_DUPLIKATE_ANZAHL[key]) + ", Dauer der doppelten Einträge: " + str(DURCHWAHL_DUPLIKATE_ZEIT[key]) + ")")
		worksheet.write(1, ROW_COUNTER_XLSX, DURCHWAHL_NAME_ORDERED[key], fett_mittig)
		worksheet.write(2, ROW_COUNTER_XLSX, key, fett_mittig)
		worksheet.write(4, ROW_COUNTER_XLSX, DURCHWAHL_ABGEHEND_ANZAHL[key],mittig)
		worksheet.write(5, ROW_COUNTER_XLSX, str(DURCHWAHL_ABGEHEND_DAUER[key]),zahlen)
		if DURCHWAHL_ABGEHEND_ANZAHL[key] > 0:
			worksheet.write(6, ROW_COUNTER_XLSX, str(datetime.timedelta(seconds=int(DURCHWAHL_ABGEHEND_DAUER[key].total_seconds()/DURCHWAHL_ABGEHEND_ANZAHL[key]))),mittig)
		else:
			worksheet.write(6, ROW_COUNTER_XLSX, "0:00:00",mittig)
		worksheet.write(8, ROW_COUNTER_XLSX, DURCHWAHL_ANKOMMEND_ANZAHL[key],mittig)
		worksheet.write(9, ROW_COUNTER_XLSX, str(DURCHWAHL_ANKOMMEND_DAUER[key]),zahlen)
		if DURCHWAHL_ANKOMMEND_ANZAHL[key] > 0:
			worksheet.write(10, ROW_COUNTER_XLSX, str(datetime.timedelta(seconds=int(DURCHWAHL_ANKOMMEND_DAUER[key].total_seconds()/DURCHWAHL_ANKOMMEND_ANZAHL[key]))),mittig)
		else:
			worksheet.write(10, ROW_COUNTER_XLSX, "0:00:00",mittig)
		worksheet.write(12, ROW_COUNTER_XLSX, str(DURCHWAHL_ABGEHEND_ANZAHL[key] + DURCHWAHL_ANKOMMEND_ANZAHL[key]),mittig)
		worksheet.write(13, ROW_COUNTER_XLSX, str(DURCHWAHL_ABGEHEND_DAUER[key] + DURCHWAHL_ANKOMMEND_DAUER[key]),mittig)
		worksheet.write(14, ROW_COUNTER_XLSX, str(datetime.timedelta(seconds=int((DURCHWAHL_ABGEHEND_DAUER[key].total_seconds() + DURCHWAHL_ANKOMMEND_DAUER[key].total_seconds())/TAGE*7))),mittig)
		ROW_COUNTER_XLSX += 1
workbook.close()
print("")
print("Auswertung abgeschlossen.")
print("Das Programm wird in wenigen Sekunden beendet.")
time.sleep(10)
sys.exit()
