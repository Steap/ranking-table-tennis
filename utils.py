# -*- coding: utf-8 -*-

import csv
import models
import os
from openpyxl import Workbook, load_workbook, worksheet

__author__ = 'sebastian'


# TODO add xlsx support to read/write multiple sheets from/to a single file
def load_league_workbook(filename):
    wb = load_workbook(filename, read_only=True)
    snames = wb.sheetnames


    print snames

    print [sname for sname in snames if "Partidos" in sname]

    ws = wb.get_sheet_by_name(snames[7])

    print ws


def load_sheet_workbook(filename, sheetname):
    wb = load_workbook(filename, read_only=True)
    ws = wb.get_sheet_by_name(sheetname)

    ws.calculate_dimension(force=True)
    print ws.dimensions

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
    return list_to_return[1:]


# TODO add support for multiple rows header
def save_csv(filename, headers, list_to_save):
    with open(filename, 'w') as outcsv:
        writer = csv.writer(outcsv)
        writer.writerow(headers)
        writer.writerows(list_to_save)


def load_csv(filename):
    with open(filename, 'r') as incsv:
        reader = csv.reader(incsv)
        aux = [row for row in reader]
        header = aux[0]
        list_to_return = []
        for row in aux[1:]:
            aux_row = []
            for item in row:
                if item.isdigit():
                    aux_row.append(int(item))
                else:
                    aux_row.append(item)
            list_to_return.append(aux_row)
        return list_to_return


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
for i, category in enumerate(["primera", "segunda", "tercera"]):
    round_points[category] = {}
    for row in aux_round_points:
        priority = row[0]
        reached_round = row[1]
        points = row[2+i]
        round_points[category][reached_round] = points
        rounds_priority[reached_round] = priority


def points_to_assign(rating_winner, rating_loser):
    """Return points to assign to each player given"""
    rating_diff = rating_winner - rating_loser

    assignation_table = expected_result
    if rating_diff < 0:
        rating_diff *= -1.0
        assignation_table = unexpected_result

    i = 0
    while rating_diff > assignation_table[i][0]:
        i += 1

    points_to_winner = assignation_table[i][1]
    points_to_loser = assignation_table[i][2]
    # print "diff:%d, to_winner:%d, to_loser:%d" % (rating_diff, points_to_winner, points_to_loser)

    return [points_to_winner, points_to_loser]


def load_ranking_csv(filename):
    # TODO add support for rankings csvs with expanded header
    # """Loads an csv and return a preprocessed ranking (name, date, ranking_list)"""
    # with open(filename, 'r') as incsv:
    #     reader = csv.reader(incsv)
    #     aux = [row for row in reader]
    #
    #     name = aux[0][1]
    #     date = aux[1][1]
    #     location = aux[2][1]
    #
    #     raw_ranking = aux[4:]
    #     ranking = models.Ranking(name, date)
    #     ranking.load_list([[r[0], r[1]] for r in raw_ranking])
    raw_ranking = load_csv(filename)
    # TODO name date and location should be read from file
    ranking_list = [[r[0], r[2], r[3]] for r in raw_ranking]
    return ranking_list


def load_tournament_csv(filename):
    """Loads an csv and return a preprocessed match list (winner, loser, round, category) and a list of players"""
    with open(filename, 'r') as incsv:
        reader = csv.reader(incsv)
        tournament_list = [row for row in reader]
        return load_tournament_list(tournament_list)


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
