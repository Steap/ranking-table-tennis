# -*- coding: utf-8 -*-


class Player:
    def __init__(self, pid=-1, name="Apellido, Nombre", association="Asociación", city="Ciudad"):
        self.pid = pid
        self.name = name
        self.association = association
        self.city = city

    def __repr__(self):
        return ";".join([str(self.pid), self.name, self.association, self.city])


class PlayersList:
    def __init__(self):
        self.players = {}

    def __getitem__(self, pid):
        return self.get_player(pid)

    def get_player(self, pid):
        return self.players.get(pid)

    def __len__(self):
        return len(self.players)

    def __repr__(self):
        return "\n".join(str(self.get_player(p)) for p in self.players)

    def add_player(self, player):
        if player.pid not in self.players:
            self.players[player.pid] = player
        else:
            print "WARNING: Already exists a player for that pid. Check:", str(player)
    
    def add_new_player(self, name):
        pid = 0
        while pid in self.players:
            pid += 1
        self.add_player(Player(pid, name))
    
    # def __contains__(self, ):

    def get_pid(self, name):
        for player in self.players.itervalues():
            if name == player.name:
                return player.pid
        print "WARNING: Unknown player:", name
        return None

    def to_list(self):
        players_list = [[p.pid, p.name, p.association, p.city] for p in self.players.itervalues()]
        return players_list

    def load_list(self, players_list):
        for pid, name, association, city in players_list:
            self.add_player(Player(pid, name, association, city))


class RankingEntry:
    def __init__(self, pid, rating, bonus):
        self.pid = pid
        self.rating = rating
        self.bonus = bonus
        
    def __repr__(self):
        return ";".join([str(self.pid), str(self.rating), str(self.bonus)])


class Ranking:
    def __init__(self):
        self.ranking = {}
        self.date = ""
        self.tournament = ""

    def add_entry(self, pid, rating, bonus):
        if pid not in self.ranking:
            self.ranking[pid] = RankingEntry(pid, rating, bonus)
        else:
            print "WARNING: Already exists an entry for pid:", pid

    def get_entry(self, pid):
        return self.ranking[pid]
        
    def __getitem__(self, pid):
        return self.get_entry(pid)

    def __repr__(self):
        return "\n".join(str(self.get_entry(re)) for re in self.ranking)





