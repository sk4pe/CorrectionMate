#!/usr/bin/env python
# -*- coding: utf-8 -*-

import xlsxwriter
from xlsxwriter.utility import xl_rowcol_to_cell
from tkinter import *
from tkinter import ttk
import tkinter
from tkinter import messagebox
import os


# Variablen
doc_name = "KA_Auswertung.xlsx"
SVERWTABLE_NAME = "SVerweisTable"
NOTENSCHLTABLE_NAME = "1. Notenschlüssel anpassen"
AUSWERTUNG_NAME = "2. Auswertung eintragen"
KOMMENTARE_NAME = "3. Kommentare eintragen"
OVERVIEW_NAME = "4. Übersicht"
RAW_NAME = "Rohdaten (Export)"

punktesystem = None
counter_table = 15
notenschl_table = None
anzahl_sus = None

row_anchor_point_table = None
col_anchor_point_table = 3
format_rot_m_b = None
format_gruen_m_b = None
format_rot_m_b_f = None
format_rot_r_b_f = None
format_gruen_m_b_f = None
format_rahmen_rot = None
format_m_f = None
format_gruen_b = None
format_rot_b = None
workbook = None

noten_eins = [
    '6',
    '5-',
    '5',
    '5+',
    '4-',
    '4',
    '4+',
    '3-',
    '3',
    '3+',
    '2-',
    '2',
    '2+',
    '1-',
    '1'
]

noten_einsplus = [
    '6',
    '5-',
    '5',
    '5+',
    '4-',
    '4',
    '4+',
    '3-',
    '3',
    '3+',
    '2-',
    '2',
    '2+',
    '1-',
    '1',
    '1+'
]

noten_textuell = noten_eins

keysSek1 = {}
keysSek2 = {}


# Methoden
def createNotenschlTable():
    global punktesystem
    global notenschl_table
    worksheet_notentable = workbook.add_worksheet(NOTENSCHLTABLE_NAME)
    worksheet_notentable.write_string(0, 0,
                                      "In diesem Blatt können die Prozentschritte für die Noten manuell angepasst werden.")
    worksheet_notentable.merge_range(2, 0, 2, 1, "BEARBEITEN", format_gruen_m_b_f)
    worksheet_notentable.merge_range(2, 3, 2, 4, "NICHT BEARBEITEN", format_rot_m_b_f)
    worksheet_notentable.merge_range(4, 0, 4, 1, "Prozent", format_rot_m_b_f)
    worksheet_notentable.write_string(4, 2, "Note", format_rot_m_b_f)
    worksheet_notentable.write_string(5, 0, "Von", format_rot_m_b_f)
    worksheet_notentable.write_string(5, 1, "Bis", format_rot_m_b_f)
    worksheet_notentable.write_string(5, 2, "", format_rot_m_b_f)



    row = 6
    col = 0

    for proz_von, proz_bis, note in notenschl_table:
        worksheet_notentable.write_formula(row, col, proz_von, format_rot_m_b)
        worksheet_notentable.write_number(row, col + 1, proz_bis, format_gruen_m_b)
        worksheet_notentable.write_string(row, col + 2, note, format_rot_m_b)
        row += 1


def createKommentare():
    worksheet_komm = workbook.add_worksheet(KOMMENTARE_NAME)
    worksheet_komm.set_column(0, 1, 16)
    worksheet_komm.set_column(2, 2, 240)

    worksheet_komm.merge_range(0, 0, 0, 2,
                               "In diesem Blatt können Kommentare für die Schülerinnen und Schüler hinterlegt werden.")
    worksheet_komm.write_string(2, 0, "Name", format_rot_b)
    worksheet_komm.write_string(2, 1, "Vorname", format_rot_b)
    worksheet_komm.write_string(2, 2, "Kommentar", format_rot_b)
    row = 3
    col = 0
    counter = 0
    while (counter < anzahl_sus):
        formula = "IF('2. Auswertung eintragen'!" + xl_rowcol_to_cell(8 + counter,
                                                                      1) + '=0,"",' + "'2. Auswertung eintragen'!" + xl_rowcol_to_cell(
            8 + counter, 1) + ")"
        worksheet_komm.write_formula(row + counter, col, formula, format_rot_b)
        formula = "IF('2. Auswertung eintragen'!" + xl_rowcol_to_cell(8 + counter,
                                                                      2) + '=0,"",' + "'2. Auswertung eintragen'!" + xl_rowcol_to_cell(
            8 + counter, 2) + ")"
        worksheet_komm.write_formula(row + counter, col + 1, formula, format_rot_b)
        worksheet_komm.write_blank(row + counter, col + 2, None, format_gruen_b)
        counter += 1


def createOverview():
    worksheet_overview = workbook.add_worksheet(OVERVIEW_NAME)
    worksheet_overview.set_column(0, 1, 16)
    worksheet_overview.merge_range(0, 0, 0, 8,
                                   "Auf diesem Blatt befindet sich eine zusammengefasste Gesamtübersicht inklusive Notenspiegel.")
    worksheet_overview.merge_range(2, 0, 2, 1, "Klasse:", format_rot_r_b_f)
    worksheet_overview.merge_range(2, 2, 2, 5, "", format_gruen_b)
    worksheet_overview.merge_range(3, 0, 3, 1, "Fach:", format_rot_r_b_f)
    worksheet_overview.merge_range(3, 2, 3, 5, "", format_gruen_b)
    worksheet_overview.merge_range(4, 0, 4, 1, "Thema:", format_rot_r_b_f)
    worksheet_overview.merge_range(4, 2, 4, 5, "", format_gruen_b)
    worksheet_overview.merge_range(5, 0, 5, 1, "Gesamtpunktzahl:", format_rot_r_b_f)
    worksheet_overview.merge_range(5, 2, 5, 5, "", format_rot_b)
    formula = "'2. Auswertung eintragen'!" + xl_rowcol_to_cell(5, 3)
    worksheet_overview.write_formula(5, 2, formula, format_rot_b)

    worksheet_overview.write_string(8, 0, "Name", format_rot_b)
    worksheet_overview.write_string(8, 1, "Vorname", format_rot_b)
    worksheet_overview.write_string(8, 2, "Summe", format_rot_b)
    worksheet_overview.write_string(8, 3, "Note", format_rot_b)

    row = 9
    col = 0
    counter = 0
    while (counter < anzahl_sus):
        formula = "=IF('2. Auswertung eintragen'!" + xl_rowcol_to_cell(8 + counter,
                                                                       1) + '=0,"",' + "'2. Auswertung eintragen'!" + xl_rowcol_to_cell(
            8 + counter, 1) + ")"
        worksheet_overview.write_formula(row + counter, col, formula, format_rot_b)
        formula = "=IF('2. Auswertung eintragen'!" + xl_rowcol_to_cell(8 + counter,
                                                                       2) + '=0,"",' + "'2. Auswertung eintragen'!" + xl_rowcol_to_cell(
            8 + counter, 2) + ")"
        worksheet_overview.write_formula(row + counter, col + 1, formula, format_rot_b)
        formula = "=IF('2. Auswertung eintragen'!" + xl_rowcol_to_cell(8 + counter,
                                                                       2 + anzahl_aufg_gesamt + 1) + '=0,"",' + "'2. Auswertung eintragen'!" + xl_rowcol_to_cell(
            8 + counter, 2 + anzahl_aufg_gesamt + 1) + ")"
        worksheet_overview.write_formula(row + counter, col + 2, formula, format_rot_b)
        formula = "=IF('2. Auswertung eintragen'!" + xl_rowcol_to_cell(8 + counter,
                                                                       2 + anzahl_aufg_gesamt + 3) + '=0,"",' + "'2. Auswertung eintragen'!" + xl_rowcol_to_cell(
            8 + counter, 2 + anzahl_aufg_gesamt + 3) + ")"
        worksheet_overview.write_formula(row + counter, col + 3, formula, format_rot_b)
        counter += 1

    # Notenspiegel
    worksheet_overview.merge_range(8, 5, 8, 6, "Notenspiegel", format_rot_m_b_f)
    for i in range(0, 6):
        worksheet_overview.write_number(9, 5 + i, i + 1, format_rot_b)
        formula = "=COUNTIF(SVerweisTable!" + xl_rowcol_to_cell(17, 1, True, True) + ":" + xl_rowcol_to_cell(
            17 + anzahl_sus, 1, True, True) + ",'4. Übersicht'!" + xl_rowcol_to_cell(9, 5 + i) + ")"
        worksheet_overview.write_formula(10, 5 + i, formula, format_rot_b)
    worksheet_overview.write_string(9, 11, "Schnitt", format_rot_b)
    formula = "=IF(SUM(" + xl_rowcol_to_cell(10, 5) + ":" + xl_rowcol_to_cell(10, 10) + ')=0,"",(' \
              + xl_rowcol_to_cell(10, 5) + "*" + xl_rowcol_to_cell(9, 5) + "+" \
              + xl_rowcol_to_cell(10, 6) + "*" + xl_rowcol_to_cell(9, 6) + "+" \
              + xl_rowcol_to_cell(10, 7) + "*" + xl_rowcol_to_cell(9, 7) + "+" \
              + xl_rowcol_to_cell(10, 8) + "*" + xl_rowcol_to_cell(9, 8) + "+" \
              + xl_rowcol_to_cell(10, 9) + "*" + xl_rowcol_to_cell(9, 9) + "+" \
              + xl_rowcol_to_cell(10, 10) + "*" + xl_rowcol_to_cell(9, 10) \
              + ")/SUM(" + xl_rowcol_to_cell(10, 5) + ":" + xl_rowcol_to_cell(10, 10) + "))"

    worksheet_overview.write_formula(10, 11, formula, format_rot_b)


def createAuswertungTable():
    worksheet_auswertungtable = workbook.add_worksheet(AUSWERTUNG_NAME)
    worksheet_auswertungtable.set_column(0, 0, 4)
    worksheet_auswertungtable.set_column(1, 2, 16)
    i = 1
    while (i <= anzahl_aufg_gesamt):
        worksheet_auswertungtable.set_column(2 + i, 2 + i, 8)
        i += 1
    worksheet_auswertungtable.merge_range(0, 0, 0, 10,
                                          "In diesem Blatt werden die Punkte der Schülerinnen und Schüler eingetragen.")
    worksheet_auswertungtable.merge_range(1, 0, 1, 10,
                                          "Zuerst muss im roten Feld die Gesamtpunktzahl eingetragen werden!")

    worksheet_auswertungtable.merge_range(3, 1, 3, 2, "BEARBEITEN", format_gruen_m_b_f)
    worksheet_auswertungtable.merge_range(4, 1, 4, 2, "NICHT BEARBEITEN", format_rot_m_b_f)
    worksheet_auswertungtable.merge_range(5, 1, 5, 2, "GESAMTPUNKTZAHL", format_m_f)
    worksheet_auswertungtable.write_number(5, 3, 0, format_rahmen_rot)

    worksheet_auswertungtable.write_string(7, 0, "", format_rot_m_b_f)
    row = 7
    col = 1
    for headline in headline_auswertung:
        worksheet_auswertungtable.write_string(row, col, headline, format_rot_m_b_f)
        col += 1

    # Schüler mit Spalten anlegen
    row = 8
    col = 0

    row_table = row_anchor_point_table
    col_table = col_anchor_point_table
    # Punkte Tafel erzeugen
    # worksheet_auswertungtable.set_row(row_anchor_point_table,30)
    worksheet_auswertungtable.write_string(row_table, col_table, "Von (%)", format_rot_m_b_f)
    worksheet_auswertungtable.write_string(row_table, col_table + 1, "Bis (%)", format_rot_m_b_f)
    worksheet_auswertungtable.write_string(row_table, col_table + 2, "Von (Punkte)", format_rot_m_b_f)
    worksheet_auswertungtable.write_string(row_table, col_table + 3, "Bis (Punkte)", format_rot_m_b_f)
    worksheet_auswertungtable.write_string(row_table, col_table + 4, "Note", format_rot_m_b_f)

    row_table += 1
    while (col_table <= 7):
        worksheet_auswertungtable.write_blank(row_table, col_table, None, format_rot_b)
        col_table += 1

    row_table += 1
    col_table = 3
    counter = 0
    anchor_point_key = 6
    while (counter < counter_table):
        formula = "='1. Notenschlüssel anpassen'!" + xl_rowcol_to_cell(anchor_point_key + counter, 0)
        worksheet_auswertungtable.write_formula(row_table + counter, col_table, formula, format_rot_m_b)

        formula = "='1. Notenschlüssel anpassen'!" + xl_rowcol_to_cell(anchor_point_key + counter, 1)
        worksheet_auswertungtable.write_formula(row_table + counter, col_table + 1, formula, format_rot_m_b)

        formula = "=Round((" + xl_rowcol_to_cell(5, 3, True, True) + "*" + xl_rowcol_to_cell(row_table + counter,
                                                                                             col_table) + ")/100,1)"
        worksheet_auswertungtable.write_formula(row_table + counter, col_table + 2, formula, format_rot_m_b)

        formula = "=Round((" + xl_rowcol_to_cell(5, 3, True, True) + "*" + xl_rowcol_to_cell(row_table + counter,
                                                                                             col_table + 1) + ")/100,1)-0.1"
        if counter == counter_table - 1:
            formula = "=Round((" + xl_rowcol_to_cell(5, 3, True, True) + "*" + xl_rowcol_to_cell(
                row_table + counter, col_table + 1) + ")/100,1)"
        worksheet_auswertungtable.write_formula(row_table + counter, col_table + 3, formula, format_rot_m_b)

        formula = "='1. Notenschlüssel anpassen'!" + xl_rowcol_to_cell(anchor_point_key + counter, 2)
        worksheet_auswertungtable.write_formula(row_table + counter, col_table + 4, formula, format_rot_m_b)
        counter += 1

    # Schüler Matrix erzeugen
    schueler = 1
    while (schueler <= anzahl_sus):
        worksheet_auswertungtable.write_number(row, col, schueler, format_rot_m_b_f)
        worksheet_auswertungtable.write_string(row, col + 1, "", format_gruen_b)
        worksheet_auswertungtable.write_string(row, col + 2, "", format_gruen_b)
        aufgabe = 1
        while (aufgabe <= anzahl_aufg_gesamt):
            worksheet_auswertungtable.write_blank(row, col + 2 + aufgabe, None, format_gruen_b)
            aufgabe += 1

        # Formel Spalte Summe
        formula = "=Sum(" + xl_rowcol_to_cell(row, 3) + ":" + xl_rowcol_to_cell(row, 3 + anzahl_aufg_gesamt - 1) + ")"
        worksheet_auswertungtable.write_formula(row, col + 2 + aufgabe, formula, format_rot_b)

        # Formel Spalte Prozent
        formula = "IF(" + xl_rowcol_to_cell(8 + anzahl_sus, col + 2 + aufgabe, True,
                                            True) + "=0,0,(" + xl_rowcol_to_cell(row,
                                                                                 3 + anzahl_aufg_gesamt) + "/" + xl_rowcol_to_cell(
            8 + anzahl_sus, col + 2 + aufgabe, True, True) + ")*100)"
        worksheet_auswertungtable.write_formula(row, col + 2 + aufgabe + 1, formula, format_rot_b)

        # Formel Spalte Note
        formula = "=IF(" + xl_rowcol_to_cell(row, col + 3) + "=" + '"","",VLOOKUP(' + xl_rowcol_to_cell(row,
                                                                                                        col + 2 + aufgabe + 1) + "," + xl_rowcol_to_cell(
            row_anchor_point_table + 2, col_anchor_point_table, True, True) + ":" + xl_rowcol_to_cell(
            row_anchor_point_table + counter_table + 2 - 1, col_anchor_point_table + 4, True, True) + ",5,TRUE))"
        worksheet_auswertungtable.write_formula(row, col + 2 + aufgabe + 2, formula, format_rot_b)
        schueler += 1
        row += 1

    # Zu erreichende Punktzahl
    worksheet_auswertungtable.merge_range(row, 1, row, 2, "Zu erreichende Punktzahl:", format_rot_b)
    aufgabe = 1
    while (aufgabe <= anzahl_aufg_gesamt):
        worksheet_auswertungtable.write_blank(row, col + 2 + aufgabe, None, format_gruen_b)
        aufgabe += 1
    formula = "=" + xl_rowcol_to_cell(5, 3)
    worksheet_auswertungtable.write_formula(row, col + 2 + aufgabe, formula, format_rot_b)
    formula = "=IF(" + xl_rowcol_to_cell(row, col + 2 + aufgabe) + "=SUM(" + xl_rowcol_to_cell(row,
                                                                                               3) + ":" + xl_rowcol_to_cell(
        row, col + 2 + aufgabe - 1) + '),"","ACHTUNG: Summe der Punkte stimmt nicht mit der Gesamtpunktzahl überein!")'
    worksheet_auswertungtable.write_formula(row, col + 2 + aufgabe + 1, formula, workbook.add_format({'bold': True}))

    # Durchschnitt
    row += 1
    worksheet_auswertungtable.merge_range(row, 1, row, 2, "Durchschnitt:", format_rot_b)
    aufgabe = 1
    while (aufgabe <= anzahl_aufg_gesamt):
        formula = "=IF(COUNT(" + xl_rowcol_to_cell(8, col + 2 + aufgabe) + ":" + xl_rowcol_to_cell(8 + anzahl_sus - 1,
                                                                                                   col + 2 + aufgabe) + ")=0,0,SUM(" + xl_rowcol_to_cell(
            8, col + 2 + aufgabe) + ":" + xl_rowcol_to_cell(8 + anzahl_sus - 1,
                                                            col + 2 + aufgabe) + ")/COUNT(" + xl_rowcol_to_cell(8,
                                                                                                                col + 2 + aufgabe) + ":" + xl_rowcol_to_cell(
            8 + anzahl_sus - 1, col + 2 + aufgabe) + "))"
        worksheet_auswertungtable.write_formula(row, col + 2 + aufgabe, formula, format_rot_b)
        aufgabe += 1

    # Durchschnitt (Prozent)
    row += 1
    worksheet_auswertungtable.merge_range(row, 1, row, 2, "Durchschnitt (Prozent):", format_rot_b)
    aufgabe = 1
    while (aufgabe <= anzahl_aufg_gesamt):
        formula = "=IF(" + xl_rowcol_to_cell(8 + anzahl_sus, col + 2 + aufgabe) + "=0,0,(" + xl_rowcol_to_cell(
            8 + anzahl_sus + 1, col + 2 + aufgabe) + "/" + xl_rowcol_to_cell(8 + anzahl_sus,
                                                                             col + 2 + aufgabe) + ")*100)"
        worksheet_auswertungtable.write_formula(row, col + 2 + aufgabe, formula, format_rot_b)
        aufgabe += 1


def createSVerweisTable():
    worksheet_sverwtable = workbook.add_worksheet(SVERWTABLE_NAME)
    notentable = (
        ["6", "ungenügend", "", 6],
        ["5-", "mangelhaft", "minus", 5],
        ["5", "mangelhaft", "", 5],
        ["5+", "mangelhaft", "plus", 5],
        ["4-", "ausreichend", "minus", 4],
        ["4", "ausreichend", "", 4],
        ["4+", "ausreichend", "plus", 4],
        ["3-", "befriedigend", "minus", 3],
        ["3", "befriedigend", "", 3],
        ["3+", "befriedigend", "plus", 3],
        ["2-", "gut", "minus", 2],
        ["2", "gut", "", 2],
        ["2+", "gut", "plus", 2],
        ["1-", "sehr gut", "minus", 1],
        ["1", "sehr gut", "", 1]

    )

    if (punktesystem.get()):
        notentable += (["1+", "sehr gut", "plus", 1],)

    row = 0
    col = 0

    for note, note_text, tendenz_text, note_num in notentable:
        worksheet_sverwtable.write_string(row, col, note, format_rot_b)
        worksheet_sverwtable.write_string(row, col + 1, note_text, format_rot_b)
        worksheet_sverwtable.write_string(row, col + 2, tendenz_text, format_rot_b)
        worksheet_sverwtable.write_number(row, col + 3, note_num, format_rot_b)
        row += 1

    row = 17
    col = 0
    counter = 0
    while (counter < anzahl_sus):
        formula = "='2. Auswertung eintragen'!" + xl_rowcol_to_cell(8 + counter, 2 + anzahl_aufg_gesamt + 3)
        worksheet_sverwtable.write_formula(row + counter, col, formula, format_rot_b)
        formula = "=IF(" + xl_rowcol_to_cell(row + counter, col) + '="","",VLOOKUP(' + xl_rowcol_to_cell(row + counter,
                                                                                                         col) + "," + xl_rowcol_to_cell(
            0, 0, True, True) + ":" + xl_rowcol_to_cell(counter_table - 1, 3, True, True) + ",4,))"
        worksheet_sverwtable.write_formula(row + counter, col + 1, formula, format_rot_b)
        counter += 1
    worksheet_sverwtable.hide()


def createRohdaten():
    global workbook
    worksheet_raw = workbook.add_worksheet(RAW_NAME)
    global anzahl_sus
    global anzahl_Unteraufgaben

    formula = "IF('2. Auswertung eintragen'!" + xl_rowcol_to_cell(7,
                                                                  1) + '=0,"",' + "'2. Auswertung eintragen'!" + xl_rowcol_to_cell(
        7, 1) + ")"
    worksheet_raw.write_formula(0, 0, formula, format_rot_b)
    formula = "IF('2. Auswertung eintragen'!" + xl_rowcol_to_cell(7,
                                                                  2) + '=0,"",' + "'2. Auswertung eintragen'!" + xl_rowcol_to_cell(
        7, 2) + ")"

    worksheet_raw.write_formula(0, 1, formula, format_rot_b)

    aufgabe = 1
    col = 2
    for anzUAufg in anzahl_Unteraufgaben:
        worksheet_raw.write_string(0, col, "A" + str(aufgabe), format_rot_b)
        col += 1
        for i in range(0, anzUAufg):
            worksheet_raw.write_string(0, col, "A" + str(aufgabe) + " " + chr(ord('a') + i) + ")", format_rot_b)
            col += 1
        aufgabe += 1

    aufgabe = 1
    for anzUAufg in anzahl_Unteraufgaben:
        worksheet_raw.write_string(0, col, "A" + str(aufgabe) + " gesamt", format_rot_b)
        col += 1
        for i in range(0, anzUAufg):
            worksheet_raw.write_string(0, col, "A" + str(aufgabe) + " " + chr(ord('a') + i) + ") gesamt", format_rot_b)
            col += 1
        aufgabe += 1

    worksheet_raw.write_string(0, col, "Summe", format_rot_b)
    col += 1

    worksheet_raw.write_string(0, col, "Summe gesamt", format_rot_b)
    col += 1
    worksheet_raw.write_string(0, col, "Prozent", format_rot_b)
    col += 1
    worksheet_raw.write_string(0, col, "Note", format_rot_b)
    col += 1
    worksheet_raw.write_string(0, col, "Note (Text)", format_rot_b)
    col += 1
    worksheet_raw.write_string(0, col, "Tendenz (Text)", format_rot_b)
    col += 1
    worksheet_raw.write_string(0, col, "Kommentar", format_rot_b)
    col += 1

    anchor_ausw_row = 8
    anchor_ausw_col = 3
    row_ausw = anchor_ausw_row
    col_ausw = anchor_ausw_col
    row = 1
    col = 0

    for i in range(0, anzahl_sus):

        # Name, Vorname füllen
        formula = "IF('2. Auswertung eintragen'!" + xl_rowcol_to_cell(7 + row,
                                                                      1) + '=0,"",' + "'2. Auswertung eintragen'!" + xl_rowcol_to_cell(
            7 + row, 1) + ")"
        worksheet_raw.write_formula(0 + row, 0, formula, format_rot_b)
        formula = "IF('2. Auswertung eintragen'!" + xl_rowcol_to_cell(7 + row,
                                                                      2) + '=0,"",' + "'2. Auswertung eintragen'!" + xl_rowcol_to_cell(
            7 + row, 2) + ")"

        worksheet_raw.write_formula(0 + row, 1, formula, format_rot_b)

        # Aufgaben Punkte füllen
        col = 2
        col_ausw = anchor_ausw_col
        for uAufg in anzahl_Unteraufgaben:
            if uAufg == 0:
                formula = "=IF('2. Auswertung eintragen'!" + xl_rowcol_to_cell(row_ausw,
                                                                               col_ausw) + '="","",' + "'2. Auswertung eintragen'!" + xl_rowcol_to_cell(
                    row_ausw, col_ausw) + ")"
                worksheet_raw.write_formula(row, col, formula, format_rot_b)
                col += 1
                col_ausw += 1
            else:
                formula = "=IF(" + xl_rowcol_to_cell(row, 0) + '="","",SUM(' + xl_rowcol_to_cell(row,
                                                                                                 col + 1) + ":" + xl_rowcol_to_cell(
                    row, col + uAufg) + "))"
                worksheet_raw.write_formula(row, col, formula, format_rot_b)
                col += 1
                for j in range(0, uAufg):
                    formula = "=IF('2. Auswertung eintragen'!" + xl_rowcol_to_cell(row_ausw,
                                                                                   col_ausw) + '="","",' + "'2. Auswertung eintragen'!" + xl_rowcol_to_cell(
                        row_ausw, col_ausw) + ")"
                    worksheet_raw.write_formula(row, col, formula, format_rot_b)
                    col += 1
                    col_ausw += 1

        col_ausw = anchor_ausw_col
        for uAufg in anzahl_Unteraufgaben:
            if uAufg == 0:
                formula = "=IF(" + xl_rowcol_to_cell(row,
                                                     0) + '="","",IF(' + "'2. Auswertung eintragen'!" + xl_rowcol_to_cell(
                    anchor_ausw_row + anzahl_sus, col_ausw, True,
                    True) + '=0,"",' + "'2. Auswertung eintragen'!" + xl_rowcol_to_cell(anchor_ausw_row + anzahl_sus,
                                                                                        col_ausw, True, True) + "))"
                worksheet_raw.write_formula(row, col, formula, format_rot_b)
                col += 1
                col_ausw += 1
            else:
                formula = "=IF(" + xl_rowcol_to_cell(row, 0) + '="","",SUM(' + xl_rowcol_to_cell(row,
                                                                                                 col + 1) + ":" + xl_rowcol_to_cell(
                    row, col + uAufg) + "))"
                worksheet_raw.write_formula(row, col, formula, format_rot_b)
                col += 1
                for j in range(0, uAufg):
                    formula = "=IF(" + xl_rowcol_to_cell(row,
                                                         0) + '="","",IF(' + "'2. Auswertung eintragen'!" + xl_rowcol_to_cell(
                        anchor_ausw_row + anzahl_sus, col_ausw, True,
                        True) + '=0,"",' + "'2. Auswertung eintragen'!" + xl_rowcol_to_cell(
                        anchor_ausw_row + anzahl_sus, col_ausw, True, True) + "))"
                    worksheet_raw.write_formula(row, col, formula, format_rot_b)
                    col += 1
                    col_ausw += 1

        # Summe
        formula = "=IF('2. Auswertung eintragen'!" + xl_rowcol_to_cell(row_ausw,
                                                                       col_ausw) + '=0,"",' + "'2. Auswertung eintragen'!" + xl_rowcol_to_cell(
            row_ausw, col_ausw) + ")"
        worksheet_raw.write_formula(row, col, formula, format_rot_b)
        col += 1
        col_ausw += 1

        # Summe ges
        formula = "=IF(" + xl_rowcol_to_cell(row, 0) + '="","",IF(' + "'2. Auswertung eintragen'!" + xl_rowcol_to_cell(
            5, 3, True, True) + '=0,"",' + "'2. Auswertung eintragen'!" + xl_rowcol_to_cell(5, 3, True, True) + "))"
        worksheet_raw.write_formula(row, col, formula, format_rot_b)
        col += 1

        # Prozent
        formula = "=IF('2. Auswertung eintragen'!" + xl_rowcol_to_cell(row_ausw,
                                                                       col_ausw) + '=0,"",' + "'2. Auswertung eintragen'!" + xl_rowcol_to_cell(
            row_ausw, col_ausw) + ")"
        worksheet_raw.write_formula(row, col, formula, format_rot_b)
        col += 1
        col_ausw += 1

        # Note
        formula = "=IF('2. Auswertung eintragen'!" + xl_rowcol_to_cell(row_ausw,
                                                                       col_ausw) + '=0,"",' + "'2. Auswertung eintragen'!" + xl_rowcol_to_cell(
            row_ausw, col_ausw) + ")"
        worksheet_raw.write_formula(row, col, formula, format_rot_b)
        col += 1
        col_ausw += 1

        # Note (text)
        formula = "=IF(" + xl_rowcol_to_cell(row, col - 1) + '="","",VLOOKUP(' + xl_rowcol_to_cell(row,
                                                                                                   col - 1) + ",SVerweisTable!" + xl_rowcol_to_cell(
            0, 0, True, True) + ":" + xl_rowcol_to_cell(counter_table - 1, 1, True, True) + ",2,FALSE))"
        worksheet_raw.write_formula(row, col, formula, format_rot_b)
        col += 1

        # Tendenz (text)
        formula = "=If(" + xl_rowcol_to_cell(row, col - 2) + '="","",IF(VLOOKUP(' + xl_rowcol_to_cell(row,
                                                                                                      col - 2) + ',SVerweisTable!' + xl_rowcol_to_cell(
            0, 0, True, True) + ":" + xl_rowcol_to_cell(counter_table - 1, 2, True,
                                                        True) + ',3,FALSE)=0,"",VLOOKUP(' + xl_rowcol_to_cell(row,
                                                                                                              col - 2) + ',SVerweisTable!' + xl_rowcol_to_cell(
            0, 0, True, True) + ":" + xl_rowcol_to_cell(counter_table - 1, 2, True, True) + ',3,FALSE)))'
        worksheet_raw.write_formula(row, col, formula, format_rot_b)
        col += 1

        # Kommentar
        formula = "=IF('3. Kommentare eintragen'!" + xl_rowcol_to_cell(row + 2,
                                                                       2) + '=0,"",' + "'3. Kommentare eintragen'!" + xl_rowcol_to_cell(
            row + 2, 2) + ")"
        worksheet_raw.write_formula(row, col, formula, format_rot_b)
        row_ausw += 1
        row += 1
    worksheet_raw.hide()


def generateDoc():
    # Datei öffnen
    global workbook
    global format_rot_m_b
    global format_gruen_m_b
    global format_rot_m_b_f
    global format_rot_r_b_f
    global format_gruen_m_b_f
    global format_rahmen_rot
    global format_m_f
    global format_gruen_b
    global format_rot_b
    global anzahl_sus
    global row_anchor_point_table
    global punktesystem
    global counter_table
    global noten_textuell
    global noten_einsplus

    if punktesystem.get():
        counter_table = 16
        noten_textuell = noten_einsplus
    workbook = xlsxwriter.Workbook(doc_name)

    row_anchor_point_table = 8 + anzahl_sus + 5
    # Formatvorlagen

    # Rot Center Border
    format_rot_m_b = workbook.add_format(
        {'align': 'center',
         'bg_color': '#f8cbad',
         'border': 1}
    )

    # Grün Center Border
    format_gruen_m_b = workbook.add_format(
        {'align': 'center',
         'bg_color': '#c6e0b4',
         'border': 1}
    )

    # Rot Center Border Fett
    format_rot_m_b_f = workbook.add_format(
        {'align': 'center',
         'bg_color': '#f8cbad',
         'border': 1,
         'bold': True}
    )
    format_rot_m_b_f.set_align('vjustify')

    # Rot Right Border Fett
    format_rot_r_b_f = workbook.add_format(
        {'align': 'right',
         'bg_color': '#f8cbad',
         'border': 1,
         'bold': True}
    )

    # Grün Center Border Fett
    format_gruen_m_b_f = workbook.add_format(
        {'align': 'center',
         'bg_color': '#c6e0b4',
         'border': 1,
         'bold': True}
    )

    # Border in Rot
    format_rahmen_rot = workbook.add_format(
        {'border': 2,
         'border_color': 'red'}
    )

    # Center Fett
    format_m_f = workbook.add_format(
        {'bold': True,
         'align': 'center'}
    )

    # Grün Border
    format_gruen_b = workbook.add_format(
        {'border': 1,
         'bg_color': '#c6e0b4'}
    )

    # Rot Border
    format_rot_b = workbook.add_format(
        {'border': 1,
         'bg_color': '#f8cbad'}
    )

    # NotenschlüsselTable anlegen
    createNotenschlTable()

    # AuswertungEintragen anlegen
    createAuswertungTable()

    # KommentareEinfügen anlegen
    createKommentare()

    # Übersicht anlegen
    createOverview()

    # SVerweisTable anlegen
    createSVerweisTable()

    # Rohdaten (Export anlegen
    createRohdaten()

    workbook.close()


# Abfrage Daten
# anz_aufg_raw = None
# while type(anz_aufg_raw) != int:
#     try:
#         anz_aufg_raw = int(input("Anzahl der Aufgaben (ohne Unteraufgaben - a), b), c), etc.):      "))
#     except ValueError:
#         print("Fehler: Bitte eine Zahl größer als 0 eingeben")
#     else:
#         if anz_aufg_raw <= 0:
#             anz_aufg_raw = None
#             print("Fehler: Bitte eine Zahl größer als 0 eingeben")
#
# anzahl_aufgaben = anz_aufg_raw
#
# print("Nun werden die Unteraufgaben abgefragt:")
# anzahl_unteraufgaben = []
# aufgabe = 1
# while(aufgabe <= anzahl_aufgaben):
#     print("--------------------------------------------------")
#     print(" ")
#     anz_unteraufg_raw = None
#     try:
#         anz_unteraufg_raw = int(input("Anzahl der Unteraufgaben in Aufgabe "+ str(aufgabe)+":        " ))
#     except ValueError:
#         print("Fehler: Bitte eine Zahl größer gleich 0 eingeben")
#     else:
#         if anz_unteraufg_raw < 0:
#             anz_unteraufg_raw = None
#             print("Fehler: Bitte eine Zahl größer gleich 0 eingeben")
#         else:
#             anzahl_unteraufgaben.append(anz_unteraufg_raw)
#             aufgabe +=1
#
#
#
# summe = 0
# array_aufgaben_text = []
# for i in range(0,len(anzahl_unteraufgaben)):
#
#     if anzahl_unteraufgaben[i] == 0:
#         array_aufgaben_text.append("A" + str(i + 1))
#         summe +=1
#     else:
#
#         for j in range(0,anzahl_unteraufgaben[i]):
#             array_aufgaben_text.append("A" + str(i+1)+" "+chr(ord('a')+j)+")")
#
#     summe += anzahl_unteraufgaben[i]
#
# anzahl_aufg_gesamt = summe
#
# headline_auswertung = [
#     'Name',
#     'Vorname'] + array_aufgaben_text+ [
#     'Summe',
#     'Prozent',
#     'Note'
# ]

root = Tk()

content = ttk.Frame(root)

root.title("Correction Mate v0.2")
root.geometry("300x600")
try:
    root.wm_iconbitmap('icon.ico')
except:
    print("Icon fehlt!")
# Titlebild
try:

    photo = tkinter.PhotoImage(file=str(os.path.join(os.path.abspath("."),"title.gif")))
    w = ttk.Label(root, image=photo)
    w.grid(column=0, row=0, columnspan=5)
except:
    print("Bild fehlt!")

frameUAufg = None
textfields = []
sendButton = None
anzahl_Unteraufgaben = []

def setKey():
    global notenschl_table

    system = keyVar.get()
    systemList = None
    if punktesystem.get() and system != '':
        systemList = keysSek2[system]
        notenschl_table = (
            ["=0", systemList[15], "6"],
            ["=B7", systemList[14], "5-"],
            ["=B8", systemList[13], "5"],
            ["=B9", systemList[12], "5+"],
            ["=B10", systemList[11], "4-"],
            ["=B11", systemList[10], "4"],
            ["=B12", systemList[9], "4+"],
            ["=B13", systemList[8], "3-"],
            ["=B14", systemList[7], "3"],
            ["=B15", systemList[6], "3+"],
            ["=B16", systemList[5], "2-"],
            ["=B17", systemList[4], "2"],
            ["=B18", systemList[3], "2+"],
            ["=B19", systemList[2], "1-"],
            ["=B20", systemList[1], "1"],
            ["=B21", systemList[0], "1+"],
        )
    elif system != '':
        systemList = keysSek1[system]
        notenschl_table = (
            ["=0", systemList[14], "6"],
            ["=B7", systemList[13], "5-"],
            ["=B8", systemList[12], "5"],
            ["=B9", systemList[11], "5+"],
            ["=B10", systemList[10], "4-"],
            ["=B11", systemList[9], "4"],
            ["=B12", systemList[8], "4+"],
            ["=B13", systemList[7], "3-"],
            ["=B14", systemList[6], "3"],
            ["=B15", systemList[5], "3+"],
            ["=B16", systemList[4], "2-"],
            ["=B17", systemList[3], "2"],
            ["=B18", systemList[2], "2+"],
            ["=B19", systemList[1], "1-"],
            ["=B20", systemList[0], "1"],

        )
    else:

        notenschl_table = (
            ["=0", 20, "6"],
            ["=B7", 28.2, "5-"],
            ["=B8", 36.6, "5"],
            ["=B9", 45, "5+"],
            ["=B10", 50, "4-"],
            ["=B11", 55, "4"],
            ["=B12", 60, "4+"],
            ["=B13", 65, "3-"],
            ["=B14", 70, "3"],
            ["=B15", 75, "3+"],
            ["=B16", 80, "2-"],
            ["=B17", 85, "2"],
            ["=B18", 90, "2+"],
            ["=B19", 95, "1-"],
            ["=B20", 100, "1"],
        )
        if punktesystem.get():
            notenschl_table = (
                ["=0", 20, "6"],
                ["=B7", 27, "5-"],
                ["=B8", 34, "5"],
                ["=B9", 40, "5+"],
                ["=B10", 45, "4-"],
                ["=B11", 50, "4"],
                ["=B12", 55, "4+"],
                ["=B13", 60, "3-"],
                ["=B14", 65, "3"],
                ["=B15", 70, "3+"],
                ["=B16", 75, "2-"],
                ["=B17", 80, "2"],
                ["=B18", 85, "2+"],
                ["=B19", 90, "1-"],
                ["=B20", 95, "1"],
                ["=B21", 100, "1+"],
            )

# Funktionen
def unteraufgabenSetzen():
    global anzahl_Unteraufgaben
    global anzahl_aufg_gesamt
    global headline_auswertung
    global root
    global entSuS
    global anzahl_sus

    ok = True
    anzahl_Unteraufgaben = []

    for tf in textfields:
        try:
            inputValue = int(tf.get())
        except ValueError:
            messagebox.showinfo(title="Fehler!", message="Anzahl der Unteraufgaben muss 0 oder größer sein!")
            ok = False
            break
        else:
            anzahl_Unteraufgaben.append(inputValue)
    if ok:

        anzahl_sus = int(entSuS.get())
        summe = 0
        array_aufgaben_text = []
        for i in range(0, len(anzahl_Unteraufgaben)):

            if anzahl_Unteraufgaben[i] == 0:
                array_aufgaben_text.append("A" + str(i + 1))
                summe += 1
            else:

                for j in range(0, anzahl_Unteraufgaben[i]):
                    array_aufgaben_text.append("A" + str(i + 1) + " " + chr(ord('a') + j) + ")")

            summe += anzahl_Unteraufgaben[i]

        anzahl_aufg_gesamt = summe

        headline_auswertung = [
                                  'Name',
                                  'Vorname'] + array_aufgaben_text + [
                                  'Summe',
                                  'Prozent',
                                  'Note'
                              ]

        setKey()
        try:

            generateDoc()
        except:
           messagebox.showinfo(title="Fehler!",message="Erzeugen der Tabelle fehlgeschlagen. Bitte die Excel-Datei schließen!")

        else:
            root.destroy()


def aufgabenSetzen(*args):
    global frameUAufg
    global textfields
    global sendButton
    global aufgVal

    if frameUAufg is not None:
        frameUAufg.grid_forget()
        frameUAufg.destroy()
        frameUAufg = None
        textfields = []
    if sendButton is not None:
        sendButton.grid_forget()
        sendButton.destroy()
        sendButton = None
    if entAufg.get() != "":

        try:
            inputValue = int(entAufg.get())

        except ValueError:
            messagebox.showinfo(title="Fehler!", message="Bitte eine Zahl größer als 0 eingeben!")
            aufgVal.set(str(1))
        else:

            if inputValue > 0:
                anzahl_aufgaben = inputValue

                frameUAufg = ttk.Labelframe(root, borderwidth=2, relief="solid", text='Anzahl Teilaufgaben eingeben:')
                frameUAufg.grid(column=0, row=6, sticky=(W, N))

                for i in range(0, anzahl_aufgaben):
                    lbl = ttk.Label(frameUAufg, text="Aufgabe " + str(i + 1) + ":")
                    lbl.grid(row=0 + i, column=0, sticky=(W, N))
                    uAufgVal = StringVar()
                    txt = Spinbox(frameUAufg, from_=0, to=10, textvariable=uAufgVal)
                    txt.grid(row=0 + i, column=1, sticky=(W, N))

                    textfields.append(uAufgVal)

                sendButton = ttk.Button(root, text="Tabelle erstellen!", command=unteraufgabenSetzen)
                sendButton.grid(column=0, row=7, sticky=(W, N))
            else:
                messagebox.showinfo(title="Fehler!", message="Bitte eine Zahl größer als 0 eingeben!")


def checkSchueler(*args):
    global entSuS
    global suSVal
    if entSuS.get() != "":
        try:
            inputValue = int(entSuS.get())
        except ValueError:
            messagebox.showinfo(title="Fehler!", message="Bitte eine Zahl größer als 0 eingeben!")
            suSVal.set(str(1))
        else:
            if int(entSuS.get()) == 0: suSVal.set(str(1))


def searchKeySystems():
    path = os.path.abspath(".")
    path = os.path.join(path ,"Notenschluessel")
    if not os.path.isdir(path):
        os.makedirs(path, exist_ok=True)
        file = open(os.path.join(path,'BeispielSek1.txt'),'w+')
        file.write('Name = BeispielSek1' + os.linesep)
        file.write('Hoechste Note = 1' + os.linesep)
        file.write('System = 100,95,90,85,80,75,70,65,60,55,50,45,36.6,28.2,20,0')
        file.close()

        file = open(os.path.join(path,'BeispielSek2.txt'), 'w+')
        file.write('Name = BeispielSek2' + os.linesep)
        file.write('Hoechste Note = 1+' + os.linesep)
        file.write('System = 100,95,90,85,80,75,70,65,60,55,50,45,40,34,27,20,0')
        file.close()

    keyList = os.listdir(path)
    print(keyList)
    for file in keyList:
        fileOkay = True
        noteEinsPlus = False
        data = []
        name = ''
        if os.path.isfile(os.path.join(path,file)):

            file = open(os.path.join(path,file),'r')

            for line in file:

                line.replace(os.linesep,"")
                line = line.lstrip().rstrip()
                if line != "":


                    if line.startswith('Hoechste Note =') and fileOkay:
                        print("1")
                        if(line.find('1+')) != -1:
                            noteEinsPlus = True
                        elif line.find('1') != -1:
                            noteEinsPlus = False
                        else:
                            fileOkay = False
                    elif line.startswith('System =') and fileOkay:
                        print("2")
                        line = line[line.index('=')+1:]
                        line = line.lstrip().rstrip()

                        while line.find(',') != -1:
                            data.append(float(line[:line.index(',')]))
                            line = line[line.index(',')+1:]

                        if noteEinsPlus and data.__len__() != 16:
                            fileOkay = False
                        if not noteEinsPlus and data.__len__() != 15:
                            fileOkay = False

                    elif line.startswith('Name = ') and fileOkay:
                        print("3")
                        name = line[line.index('=')+2:]

                        name = name.lstrip().rstrip()
                        if name == '':
                            fileOkay = False
                    else:
                        fileOkay = False

            if fileOkay:
                if not noteEinsPlus:
                    keysSek1[name]=data

                else:
                    keysSek2[name]=data
            file.close()
def updateKey(*args):
    if punktesystem.get():
        keyList = ()
        for a in keysSek2.keys():
            keyList = keyList + (a,)
        comboKey['values']= keyList
        if keyList.__len__()!=0:
            comboKey.current(0)
    else:
        keyList = ()
        for a in keysSek1.keys():
            keyList = keyList+(a,)
        comboKey['values']=keyList
        if keyList.__len__()!=0:
            comboKey.current(0)

# Frame Aufgaben
aufgFrame = ttk.Labelframe(root, borderwidth=2, relief="solid", text="Anzahl Aufgaben (ohne Teilaufgaben):")
aufgVal = StringVar()
entAufg = Spinbox(aufgFrame, from_=1, to=40, textvariable=aufgVal)
aufgFrame.grid(column=0, row=1, sticky=(W, N))
entAufg.grid(column=0, row=0, sticky=(N, W))

# Frame SuS
susFrame = ttk.Labelframe(root, borderwidth=2, relief="solid", text="Anzahl der Schülerinnen und Schüler:")
suSVal = StringVar()
entSuS = Spinbox(susFrame, from_=1, to=35, textvariable=suSVal)
susFrame.grid(column=0, row=2, sticky=(W, N))
entSuS.grid(column=0, row=0, sticky=(W, N))

searchKeySystems()
# Frame Notenschlüssel
nkFrame = ttk.Labelframe(root, borderwidth=2, relief="solid", text="Notenschlüssel:")
keyVar = StringVar()
comboKey = ttk.Combobox(nkFrame, textvariable=keyVar, state='readonly')
keyList = ()
for a in keysSek1.keys():
    keyList = keyList + (a,)
comboKey['values'] = keyList
if keyList.__len__()!=0:
    comboKey.current(0)
comboKey.grid(column=0, row=0, sticky=(N, W))
nkFrame.grid(column=0, row=4, sticky=(N, W))

# Frame mit Radiobuttons
radioFrame = ttk.Labelframe(root, borderwidth=2, relief="solid", text="Hoechste Note:")
punktesystem = BooleanVar()
eins = ttk.Radiobutton(radioFrame, text="1", variable=punktesystem, value=False)
einsplus = ttk.Radiobutton(radioFrame, text="1+", variable=punktesystem, value=True)
radioFrame.grid(column=0, row=3, sticky=(W, N))
eins.grid(column=0, row=0, sticky=(W, N))
einsplus.grid(column=1, row=0, sticky=(W, N))

content.grid(column=0, row=1)

punktesystem.trace('w', updateKey)
aufgVal.trace("w", aufgabenSetzen)
suSVal.trace("w", checkSchueler)

root.resizable(False, True)
root.mainloop()
