import csv
import models
import os
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment

__author__ = 'sebastian'


def get_sheetnames_by_date(filename, filter_key=""):
    wb = load_workbook(filename, read_only=True)
    sheetnames = [s for s in wb.sheetnames if filter_key in s]
    namesdates = [(name, load_tournament_xlsx(filename, name)["date"]) for name in sheetnames]
    namesdates.sort(key=lambda p: p[1])

    return [name for name, date in namesdates]


def load_sheet_workbook(filename, sheetname, first_row=1):
    wb = load_workbook(filename, read_only=True)
    ws = wb.get_sheet_by_name(sheetname)

    ws.calculate_dimension(force=True)
    # print(ws.dimensions)

    list_to_return = []
    # max_column = 0
    for row in ws.rows:
        # if not max_column:
        #     max_column = 1
        aux_row = []
        for cell in row:
            if cell.value is None:
                aux_row.append("")
            else:
                aux_row.append(cell.value)
        list_to_return.append(aux_row)
    return list_to_return[first_row:]


def save_sheet_workbook(filename, sheetname, headers, list_to_save, overwrite=False):
    if os.path.isfile(filename):
        wb = load_workbook(filename)
        # TODO if I want to append info to a sheet, currently I cannot => split and
        if overwrite and sheetname in wb:
            wb.remove_sheet(wb.get_sheet_by_name(sheetname))
        ws = wb.create_sheet()
    else:
        wb = Workbook()
        ws = wb.active

    ws.title = sheetname

    ws.append(headers)
    for col in range(1, ws.max_column+1):
        cell = ws.cell(column=col, row=1)
        cell.font = Font(bold=True)

    for row in list_to_save:
        ws.append(row)

    wb.save(filename)


def save_csv(filename, headers, list_to_save):
    with open(filename, 'w') as outcsv:
        writer = csv.writer(outcsv)
        writer.writerow(headers)
        writer.writerows(list_to_save)


def load_csv(filename, first_row=1):
    with open(filename, 'r') as incsv:
        reader = csv.reader(incsv)
        list_to_return = []
        for row in reader:
            aux_row = []
            for item in row:
                if item.isdigit():
                    aux_row.append(int(item))
                else:
                    aux_row.append(item)
            list_to_return.append(aux_row)
        return list_to_return[first_row:]


##############################
# Tables to assign points
##############################

# Expected result table
# TODO, falta chequear
config_folder = os.path.dirname(__file__) + "/config/"
# difference, points to winner, points to loser
expected_result = load_csv(config_folder + "expected_result.csv")

# negative difference, points to winner, points to loser
unexpected_result = load_csv(config_folder + "unexpected_result.csv")

# points to be assigned by round
aux_round_points = load_csv(config_folder + "puntos_por_ronda.csv")
round_points = {}
rounds_priority = {}
for i, categ in enumerate(["primera", "segunda", "tercera"]):
    round_points[categ] = {}
    for r in aux_round_points:
        priority = r[0]
        reached_round = r[1]
        points = r[2 + i]
        round_points[categ][reached_round] = points
        rounds_priority[reached_round] = priority


def load_ranking_csv(filename):
    raw_ranking = load_csv(filename)
    # TODO name date and location should be read from file
    ranking_list = [[rr[0], rr[2], rr[3]] for rr in raw_ranking]
    return ranking_list


def save_ranking_sheet(filename, sheetname, ranking, players, overwrite=False):
    if os.path.isfile(filename):
        wb = load_workbook(filename)
        if overwrite and sheetname in wb:
            wb.remove_sheet(wb.get_sheet_by_name(sheetname))
        ws = wb.create_sheet()
    else:
        wb = Workbook()
        ws = wb.active

    ws.title = sheetname

    ws["A1"] = "Nombre del torneo"
    ws["B1"] = ranking.tournament_name
    ws.merge_cells('B1:G1')
    ws["A2"] = "Fecha"
    ws["B2"] = ranking.date
    ws.merge_cells('B2:G2')
    ws["A3"] = "Lugar"
    ws["B3"] = ranking.location
    ws.merge_cells('B3:G3')

    ws.append(["PID", "Total puntos", "Nivel de juego", "Puntos bonus", "Jugador", "Asociación", "Ciudad"])

    to_bold = ["A1", "A2", "A3",
               "A4", "B4", "C4", "D4", "E4", "F4", "G4"]
    to_center = to_bold + ["B1", "B2", "B3"]

    for colrow in to_bold:
        cell = ws.cell(colrow)
        cell.font = Font(bold=True)
    for colrow in to_center:
        cell = ws.cell(colrow)
        # TODO add width adaptation instead of shrink
        cell.alignment = Alignment(horizontal='center')

    list_to_save = [[e.pid, e.get_total(), e.rating, e.bonus, players[e.pid].name, players[e.pid].association,
                     players[e.pid].city] for e in ranking]

    for row in sorted(list_to_save, key=lambda l: l[1], reverse=True):
        ws.append(row)

    wb.save(filename)


def load_ranking_sheet(filename, sheetname):
    # """Loads an csv and return a preprocessed ranking (name, date, ranking_list)"""
    # TODO check if date is being read properly
    raw_ranking = load_sheet_workbook(filename, sheetname, first_row=0)
    ranking = models.Ranking(raw_ranking[0][1], raw_ranking[1][1], raw_ranking[2][1])
    ranking.load_list([[rr[0], rr[2], rr[3]] for rr in raw_ranking[4:]])
    return ranking


def load_tournament_csv(filename):
    """Loads an csv and return a preprocessed match list (winner, loser, round, category) and a list of players"""
    with open(filename, 'r') as incsv:
        reader = csv.reader(incsv)
        tournament_list = [row for row in reader]
        return load_tournament_list(tournament_list)


def load_tournament_xlsx(filename, sheet_name):
    """Loads an xlsx sheet and return a preprocessed match list (winner, loser, round, category)
    and a list of players"""
    return load_tournament_list(load_sheet_workbook(filename, sheet_name, 0))


def load_tournament_list(tournament_list):
    name = tournament_list[0][1]
    date = tournament_list[1][1]
    location = tournament_list[2][1]

    # Processing matches
    raw_match_list = tournament_list[5:]

    # Ordered list of the players of the tournament
    players_list = set()
    for row in raw_match_list:
        players_list.add(row[0])
        players_list.add(row[1])
    players_list = list(players_list)
    players_list.sort()

    # Reformated list of matches
    matches_list = []
    for player1, player2, set1, set2, round_match, category in raw_match_list:
        if int(set1) > int(set2):
            winner = player1
            loser = player2
        elif int(set1) < int(set2):
            winner = player2
            loser = player1
        else:
            print("Error al procesar los partidos, se encontró un empate entre %s y %s" % (player1, player2))
            break
        matches_list.append([winner, loser, round_match, category])

    tournament = {"name": name,
                  "date": date,
                  "location": location,
                  "players": players_list,
                  "matches": matches_list}

    return tournament
