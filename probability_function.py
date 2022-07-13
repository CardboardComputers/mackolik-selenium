import probability_machine
import os
import time

while True:
    print('░▒▓ Ev Sahibi Takım')
    ht_matches_played = int(input('Oynadığı maç sayısı: '))
    ht_goals_scored = int(input('Attığı gol: '))
    ht_goals_lost = int(input('Yediği gol: '))
    print('░▒▓ Konuk Takım')
    at_matches_played = int(input('Oynadığı maç sayısı: '))
    at_goals_scored = int(input('Attığı gol: '))
    at_goals_lost = int(input('Yediği gol: '))
    print('░▒▓ Lig Ortalaması')
    league_matches_played = int(input('Oynanan maç sayısı: '))
    league_home_goals = int(input('Ev sahibi gol: '))
    league_away_goals = int(input('Misafir gol: '))

    print('\n░▒▓ Olasılıklar:')

    cprob = probability_machine.write_spreadsheet(
        'probabilities.xlsx',
        ht_matches_played,
        ht_goals_scored,
        ht_goals_lost,
        at_matches_played,
        at_goals_scored,
        at_goals_lost,
        league_matches_played * 2,
        league_matches_played,
        league_home_goals,
        league_away_goals
    )
    print('OK, ./probabilities.xlsx\n')
    time.sleep(1)
    os.startfile('probabilities.xlsx')
