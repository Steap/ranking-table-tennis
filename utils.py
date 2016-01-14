__author__ = 'sebastian'


class Player:
    """
    Model for a player
    """

    def __init__(self, nid, name, rating=0.0):
        self.nid = nid
        self.name = name
        self.email = "user@host.com"
        self.rating = rating
        self.association = "Association"

    def __str__(self):
        return "Player nid={0:d}, name={1:s}, rating={2:d}".format(self.nid, self.name, self.rating)

    def __repr__(self):
        return "<Player nid={0:d}, name={1:s}, rating={2:d}>".format(self.nid, self.name, self.rating)


class Match:
    """
    Model for a match, were there is a winner and a loser player
    """
    def __init__(self):
        self.winner = -1
        self.looser = -1

##############################
# Tables to assign points
##############################

# Expected result table
# TODO, falta chequear
# difference, points to winner, points to loser
expected_result = [[24, 9, 9],
                   [49, 8, 8],
                   [99, 7, 7],
                   [149, 6, 6],
                   [199, 5, 5],
                   [299, 4, 4],
                   [399, 3, 3],
                   [499, 2, 2],
                   [999, 1, 1]]

# TODO, falta chequear
# difference (should be negative), points to winner, points to loser
unexpected_result = [[24, 10, 9],
                     [49, 11, 10],
                     [99, 13, 11],
                     [149, 15, 12],
                     [199, 18, 14],
                     [299, 21, 16],
                     [399, 24, 18],
                     [499, 28, 21],
                     [999, 32, 25]]

# points to be assigned by round
# TODO, falta chequear
round_points = {'z': 1,
                'o': 2,
                'q': 4,
                's': 8,
                '2': 10,
                '1': 12}


def points_to_assign(rating_winner, rating_loser):
    """Return points to assign to each player rating given """
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
    print "diff:%d, to_winner:%d, to_loser:%d" % (rating_diff, points_to_winner, points_to_loser)

    return [points_to_winner, points_to_loser]

# TODO ranking model, get_rating, update_rating
def get_rating(player_id, ranking):
    return ranking[player_id]

def get_new_ranking(old_ranking, matches_list):
    # TODO make a better way to copy a list, maybe the db saves me
    new_ranking = [list(item) for item in old_ranking]
    # TODO add new players

    for winner, loser, unused in matches_list:
        # TODO change the way to read old ratings
        [to_winner, to_loser] = points_to_assign(old_ranking[winner][1], old_ranking[loser][1])
        # TODO change the way to read new ratings
        new_ranking[winner][1] += to_winner
        new_ranking[loser][1] -= to_loser

    # TODO add points per best round reached
    best_round_to_assign = {}

    for winner, loser, round_match in matches_list:
        if best_round_to_assign.get(winner):
            if best_round_to_assign.get(winner) < round_points[round_match]:
                best_round_to_assign[winner] = round_points[round_match]
        else:
            best_round_to_assign[winner] = round_points[round_match]
        if best_round_to_assign.get(loser):
            if best_round_to_assign.get(loser) < round_points[round_match]:
                best_round_to_assign[loser] = round_points[round_match]
        else:
            best_round_to_assign[loser] = round_points[round_match]

    for player_id in best_round_to_assign:
        new_ranking[player_id][1] += best_round_to_assign[player_id]

    return new_ranking

