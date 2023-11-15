#Import
import random
from SheetSetup import GeneralSheet, PlayerSheet

MAX_NUM_OF_ITEMS = 218
TOTAL_COLS_PLAYER_CARD = 20
TOTAL_ROWS_PLAYER_CARD = 14

class PitchCallingCard:
    '''
    PitchCallingCard sets up values to print to both the player and coach card
    '''
    
    def create_pitching_list(self):
        pitching_dict = {'FI': 48, "FO": 48, "FU": 5, "BF": 2,
                        "CI": 20, "CO": 20, "RI": 20, "RO": 20,
                        "PO": 2, "P1": 3, "P2": 3, "P3": 3, "P4": 3, "P5": 2, "P6": 2, "P7": 2, "P8": 1,
                        "W": 2, "B1": 2, "B3": 1, "H": 1, "EA": 1, "EP": 2, "HD": 2, "31": 2
                        }
        pitch_list = []

        for keys in pitching_dict:
            for i in range(pitching_dict[keys]):
                pitch_list.append(keys)

        #Randomize the list
        pitch_list = random.sample(pitch_list, len(pitch_list)) 
        
        #HAS TO BE LESS THAN 218 values
        if len(pitch_list) < MAX_NUM_OF_ITEMS:
            return pitch_list
        
        raise ValueError(f'Must have fewer than {MAX_NUM_OF_ITEMS} values')  

    def print_to_player_card(self, workbook):
        pitch_list = PitchCallingCard().create_pitching_list()
        ####  PRINT VALUES TO FIRST TABLE ###
        column = 2
        count = 1
        for i in range(1, TOTAL_ROWS_PLAYER_CARD):
            for row in range(2, TOTAL_COLS_PLAYER_CARD):
                # Skip the middle row that handles keys 30 through 55
                if i == 7:
                    pass
                else:
                    workbook['Player'].cell(column, row).value = pitch_list[count]
                    count += 1
            column += 1

        # Color fill in values
        PlayerSheet().add_player_table_color(workbook)

    def create_coach_card_lists(self, player_worksheet):
        pitch_dict = {'FI': [], "FO": [], "FU": [], "BF": [], "CI": [], "CO": [], "RI": [], "RO": [],
                        "PO": [], "P1": [], "P2": [], "P3": [], "P4": [], "P5": [], "P6": [], "P7": [], "P8": [],
                        "W": [], "B1": [], "B3": [], "H": [], "EA": [], "EP": [], "HD": [], "31": []
                        }
        #Go through player card and pull numbers for corresponding lists in dictionary
        for pitch in pitch_dict:
            for i in range(1,20): #for the top half of the player card
                for j in range(2, 8):
                    if str(player_worksheet.cell(j, i ).value) == pitch:
                        pitch_dict[pitch].append(str(player_worksheet.cell(1, i).value)
                                                + str(player_worksheet.cell(j, 1).value))
            for i in range(1,20): #for the bottom half of the player card
                for j in range(9,15):
                    if str(player_worksheet.cell(j, i ).value) == pitch:
                        pitch_dict[pitch].append(str(player_worksheet.cell(8, i).value)
                                                + str(player_worksheet.cell(j, 1).value))
        return pitch_dict

    def print_to_coach_card(self, workbook):
        pitch_dict = PitchCallingCard().create_coach_card_lists(workbook['Player'])
        coachWorkSheet = workbook['Coach']
        for pitches in pitch_dict:
            for i in range(1, 29): #FIND CORRECT PITCH Key
                for j in range(1, 14):
                    if coachWorkSheet.cell(i, j).value == pitches:
                        for nums in pitch_dict[pitches]: #Print out numbers
                            coachWorkSheet.cell(i + pitch_dict[pitches].index(nums) + 1, j).value = nums

class CreateCard:
    '''
    CreateCard creates both the card for the player and the coach
    '''
    def __init__(self, file_path_name):
        self.file_path_name = file_path_name

    def create_player_and_coach_card(self):
        # Open Workbook and Setup player and coach sheets
        workbook = GeneralSheet().open_workbook(self.file_path_name)
        GeneralSheet().generateAllWorkbookFormatting(workbook, self.file_path_name)

        # Print values to player and coach sheets
        PitchCallingCard().print_to_player_card(workbook)
        PitchCallingCard().print_to_coach_card(workbook)
        GeneralSheet().save_workbook(workbook, self.file_path_name)


if __name__ == '__main__':
    FILE_NAME = 'FILE_PATH'
    CreateCard(FILE_NAME).create_player_and_coach_card()
