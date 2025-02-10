import requests
import xlsxwriter
import pandas as pd
import numpy as np
import openpyxl
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Border, Side, Font, Alignment, PatternFill, numbers

# Read API key
f = open('API_KEY.txt', 'r')
API_KEY = f.read()
f.close()

# API parameters
SPORT = 'upcoming'
REGIONS = 'us,uk,eu,au'  # uk | us | eu | au. Multiple can be specified if comma delimited
MARKETS = 'h2h,spreads,totals'  # h2h | spreads | totals. Multiple can be specified if comma delimited
ODDS_FORMAT = 'decimal'  # decimal | american
DATE_FORMAT = 'iso'  # iso | unix
BET_SIZE = 100

# Fetch odds data
odds_response = requests.get(
    f'https://api.the-odds-api.com/v4/sports/{SPORT}/odds',
    params={
        'api_key': API_KEY,
        'regions': REGIONS,
        'markets': MARKETS,
        'oddsFormat': ODDS_FORMAT,
        'dateFormat': DATE_FORMAT,
    }
)

# Check if the API request was successful
if odds_response.status_code != 200:
    print(f"Error: Unable to fetch data. Status code: {odds_response.status_code}")
    print(f"Response: {odds_response.text}")
    exit()

# Parse the JSON response
try:
    odds_response = odds_response.json()
except ValueError as e:
    print(f"Error: Unable to parse JSON response. Details: {e}")
    print(f"Response: {odds_response.text}")
    exit()

# Debugging: Print the API response
# print("API Response:", odds_response)

# Constants for indexing
BOOKMAKER_INDEX = 0
NAME_INDEX = 1
ODDS_INDEX = 2
FIRST = 0

# Event class
class Event:
    def __init__(self, data):
        self.data = data
        try:
            self.sport_key = data['sport_key']
            self.id = data['id']
        except KeyError as e:
            print(f"Error: Missing key in event data: {e}")
            print(f"Event data: {data}")
            self.sport_key = None
            self.id = None

    def find_best_odds(self):
        if not self.data.get('bookmakers'):
            print(f"No bookmakers found for event: {self.id}")
            return []

        try:
            num_outcomes = len(self.data['bookmakers'][FIRST]['markets'][FIRST]['outcomes'])
        except (KeyError, IndexError):
            print(f"Unable to determine outcomes for event: {self.id}")
            return []

        self.num_outcomes = num_outcomes
        best_odds = [[None, None, float('-inf')] for _ in range(num_outcomes)]

        bookmakers = self.data['bookmakers']
        for bookmaker in bookmakers:
            if not bookmaker.get('markets'):
                continue

            try:
                market = bookmaker['markets'][FIRST]
                if not market.get('outcomes'):
                    continue
            except (KeyError, IndexError):
                continue

            for outcome in range(num_outcomes):
                try:
                    bookmaker_odds = float(market['outcomes'][outcome]['price'])
                    current_best_odds = best_odds[outcome][ODDS_INDEX]

                    if bookmaker_odds > current_best_odds:
                        best_odds[outcome][BOOKMAKER_INDEX] = bookmaker['title']
                        best_odds[outcome][NAME_INDEX] = market['outcomes'][outcome]['name']
                        best_odds[outcome][ODDS_INDEX] = bookmaker_odds
                except (KeyError, IndexError):
                    continue

        self.best_odds = best_odds
        return best_odds

    def arbitrage(self):
        total_arbitrage_percentage = 0
        for odds in self.best_odds:
            total_arbitrage_percentage += (1.0 / odds[ODDS_INDEX])

        self.total_arbitrage_percentage = total_arbitrage_percentage
        self.expected_earnings = (BET_SIZE / total_arbitrage_percentage) - BET_SIZE

        if total_arbitrage_percentage < 1:
            return True
        return False

    def convert_decimal_to_american(self):
        best_odds = self.best_odds
        for odds in best_odds:
            decimal = odds[ODDS_INDEX]
            if decimal >= 2:
                american = (decimal - 1) * 100
            elif decimal < 2:
                american = -100 / (decimal - 1)
            odds[ODDS_INDEX] = round(american, 2)
        return best_odds

    def calculate_arbitrage_bets(self):
        bet_amounts = []
        for outcome in range(self.num_outcomes):
            individual_arbitrage_percentage = 1 / self.best_odds[outcome][ODDS_INDEX]
            bet_amount = (BET_SIZE * individual_arbitrage_percentage) / self.total_arbitrage_percentage
            bet_amounts.append(round(bet_amount, 2))

        self.bet_amounts = bet_amounts
        return bet_amounts

# Process events
events = []
for data in odds_response:
    try:
        event = Event(data)
        events.append(event)
    except Exception as e:
        print(f"Error creating event: {e}")
        print(f"Event data: {data}")

# Find arbitrage opportunities
arbitrage_events = []
for event in events:
    best_odds = event.find_best_odds()
    if event.arbitrage():
        arbitrage_events.append(event)

# Calculate arbitrage bets and convert odds
for event in arbitrage_events:
    event.calculate_arbitrage_bets()
    event.convert_decimal_to_american()

# Prepare data for Excel
MAX_OUTCOMES = max([event.num_outcomes for event in arbitrage_events], default=0)
ARBITRAGE_EVENTS_COUNT = len(arbitrage_events)

my_columns = ['ID', 'Sport Key', 'Expected Earnings'] + list(np.array([[f'Bookmaker #{outcome}', f'Name #{outcome}', f'Odds #{outcome}', f'Amount to Buy #{outcome}'] for outcome in range(1, MAX_OUTCOMES + 1)]).flatten())
dataframe = pd.DataFrame(columns=my_columns)

for event in arbitrage_events:
    row = []
    row.append(event.id)
    row.append(event.sport_key)
    row.append(round(event.expected_earnings, 2))
    for index, outcome in enumerate(event.best_odds):
        row.append(outcome[BOOKMAKER_INDEX])
        row.append(outcome[NAME_INDEX])
        row.append(outcome[ODDS_INDEX])
        row.append(event.bet_amounts[index])
    while len(row) < len(dataframe.columns):
        row.append('N/A')
    dataframe.loc[len(dataframe.index)] = row

# Save to Excel
writer = pd.ExcelWriter('bets.xlsx')
dataframe.to_excel(writer, index=False)
writer.close()