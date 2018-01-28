import xlrd


class Reader:
    __filename = ""
    __book = None
    __overview_sheet = None
    __hero_sheet = None
    __global_hero_statistics_begin_row = 44
    __combat_statistics_col_occupy = 4

    def __init__(self, filename):
        self.__filename = filename
        self.__book = xlrd.open_workbook(filename)
        self.__overview_sheet = self.__book.sheet_by_index(0)
        self.__hero_sheet = self.__book.sheet_by_index(1)

    def read_ow_data(self):
        data = {}
        data['combat'] = self.read_combat_statistics()
        data['hero_and_player'] = self.read_hero_and_player_statistics()
        data['player_detail'] = self.read_hero_statistics_of_player()
        data['hero_detail'] = self.read_hero_detail()
        return data

    def read_combat_statistics(self):
        data = []
        cols_occupy = self.__combat_statistics_col_occupy
        sheet = self.__overview_sheet
        cols_num = (sheet.ncols // cols_occupy) * cols_occupy

        for col_begin in range(0, cols_num, cols_occupy):
            item = _read_battle_statistics(sheet, col_begin)
            if item is None:
                continue
            if item.get('Round') == 'ALL':
                data.append(item)
                break
            if None not in item.values():
                data.append(item)
        return data

    def read_hero_and_player_statistics(self):
        row = self.__global_hero_statistics_begin_row
        sheet = self.__overview_sheet

        data = {}
        t1, row = _read_row_until_reach_invalid_cell(sheet, row, 0,
                                                     _read_single_hero_global_statistics)

        row = _next_valid_row_cell(sheet, row, 0) + 1
        t2, row = _read_row_until_reach_invalid_cell(sheet, row, 0,
                                                     _read_single_player_global_statistics)

        data['Global_Hero_Data'] = t1
        data['Global_Player_Data'] = t2

        hero_and_player_data = []
        while True:
            row = _next_valid_row_cell(sheet, row, 0)
            if row is None:
                break

            t1, row = _read_row_until_reach_invalid_cell(sheet, row + 1, 0,
                                                         _read_single_battle_player_statistics)
            row = _next_valid_row_cell(sheet, row, 0)
            if row is None:
                break

            t2, row = _read_row_until_reach_invalid_cell(sheet, row + 1, 0,
                                                         _read_single_battle_hero_statistics)
            hero_and_player_data.append(dict(Player_Data=t1, Hero_Data=t2))

        data['Hero_And_Player_Data'] = hero_and_player_data
        return data

    def read_hero_statistics_of_player(self):
        sheet = self.__hero_sheet
        row = _next_valid_row_cell(sheet, -1, 0)
        data = []
        while row is not None:
            t1 = _read_hero_sheet_player_info(sheet, row)
            if t1['Player_Name'] == 0:
                break
            t2, row = _read_row_until_reach_invalid_cell(sheet, row + 2, 0,
                                                         _read_hero_sheet_hero_statistics)
            t2 = list(filter(lambda h: h['Playing_Time'] > 0, t2))
            row = _next_valid_row_cell(sheet, row, 0)
            data.append(dict(Player=t1, Hero=t2))

        return data

    def read_hero_detail(self):
        row = 1
        data = []
        for i in range(2, self.__book.nsheets):
            sheet = self.__book.sheet_by_index(i)
            if sheet.nrows == 0 or sheet.ncols == 0:
                break

            col = 1
            hd = {}
            while _is_valid_cell(sheet, row, col):
                hero = cell_to_value(sheet, row, col)
                d = _read_hero_detail(sheet, row + 1, col)
                hd[hero] = d
                col = col + 1

            data.append(hd)
        return data


def _read_hero_detail(sheet, row, col):
    data = dict()
    d = _get_next_row_data(sheet, col, row, sheet.nrows)
    data['Direct_Damage_Ultimate_Associated_Key'] = next(d)
    data['TeamA_Ultimate_Using_Times'] = next(d)
    data['TeamA_Ultimate_Cause_To_Kill_Times'] = next(d)
    data['TeamA_Go_Back_Home_Times'] = next(d)
    data['TeamA_Ultimate_Missing_Times'] = next(d)
    data['TeamB_Ultimate_Using_Times'] = next(d)
    data['TeamB_Ultimate_Cause_To_Kill_Times'] = next(d)
    data['TeamB_Go_Back_Home_Times'] = next(d)
    data['TeamB_Ultimate_Missing_Times'] = next(d)

    for i in range(12):
        data['Player_{0}_Ultimate_Using_Times'.format(i + 1)] = next(d)
    for i in range(12):
        data['Player_{0}_Ultimate_Cause_To_Kill_Times'.format(i + 1)] = next(d)
    for i in range(12):
        data['Player_{0}_Die_During_Ultimate_Times'.format(i + 1)] = next(d)
    for i in range(12):
        data['Player_{0}_Killed_By_This_Hero_Times'.format(i + 1)] = next(d)

    for n in range(6):
        for i in range(12):
            data['Player_{0}_Ultimate_Kill_{1}_Times'.format(i + 1, n + 1)] = next(d)

    for i in range(12):
        data['Player_{0}_Total_Ultimate_Kills'.format(i + 1)] = next(d)

    for i in range(12):
        data['Player_{0}_Time_Spent_In_Using_This_Hero'.format(i + 1)] = next(d)

    for i in range(12):
        data['Player_{0}_Direct_Kill_Support_Times'.format(i + 1)] = next(d)

    for i in range(12):
        data['Player_{0}_Direct_Kill_Tank_Times'.format(i + 1)] = next(d)

    for i in range(12):
        data['Player_{0}_Direct_Kill_Fire_Times'.format(i + 1)] = next(d)

    for n in range(20):
        for i in range(12):
            data['Player_{0}_Kills_In_Order_{1}_Ultimate'.format(i + 1, n + 1)] = next(d)

    return data


def _read_row_until_reach_invalid_cell(sheet, row, col, read_callback):
    d = []
    while _is_valid_cell(sheet, row, col):
        d.append(read_callback(sheet, row))
        row = row + 1
    return d, row


def _read_hero_sheet_player_info(sheet, row):
    data = dict()
    data['Player_Name'] = cell_to_value(sheet, row, 1)
    data['Player_Id'] = cell_to_value(sheet, row, 3)
    data['Playing_Time'] = cell_to_value(sheet, row, 9)
    return data


def _read_hero_sheet_hero_statistics(sheet, row):
    data = dict()
    d = _get_next_col_data(sheet, row, 0, sheet.ncols)
    data['Hero_Name'] = next(d)
    data['Hero_Id'] = next(d)
    data['Playing_Time'] = next(d)
    data['Ratio_Of_Playing_Time_To_Player_Time'] = next(d)
    data['Kills'] = next(d)
    data['Deaths'] = next(d)
    data['Team_Fight_Num_Joined'] = next(d)
    data['Team_Fights_Victory_Times'] = next(d)
    data['Team_Fight_Num_With_Ultimate'] = next(d)
    data['Team_Fight_Succeed_Num_Without_Ultimate'] = next(d)
    data['Team_Fight_Succeed_Num_With_Ultimate'] = next(d)
    data['Died_During_Team_Fight_Times'] = next(d)
    data['Kill_During_Team_Fight_Times'] = next(d)
    data['Ultimate_Kills'] = next(d)
    data['Instant_Kills'] = next(d)
    data['Steals_Num'] = next(d)
    data['Be_Stolen_Num'] = next(d)
    return data


def _read_single_battle_hero_statistics(sheet, row):
    data = dict()
    d = _get_next_col_data(sheet, row, 0, sheet.ncols)
    data['Hero_Id'] = next(d)
    data['Hero_Name'] = next(d)
    data['Kills'] = next(d)
    data['First_Kills'] = next(d)
    data['Deaths'] = next(d)
    data['First_Deaths'] = next(d)
    data['Difference'] = next(d)
    data['Ultimate_Num'] = next(d)
    data['Ultimate_Kills'] = next(d)
    data['Ultimate_Efficiency'] = next(d)
    data['Playing_Time'] = next(d)
    data['Ratio_Of_Playing_Time_To_Whole_Time'] = next(d)
    data['Num_Go_Back_Home'] = next(d)
    data['Ultimate_Using_Times_In_Failed_Team_Fights'] = next(d)
    data['Ultimate_Using_Times_In_Succeed_Team_Fights'] = next(d)
    data['Team_Fight_Failed_Num_Without_Ultimate'] = next(d)
    data['Team_Fight_Succeed_Num_Without_Ultimate'] = next(d)
    return data


def _read_single_battle_player_statistics(sheet, row):
    data = dict()
    d = _get_next_col_data(sheet, row, 0, sheet.ncols)
    data['Player_Name'] = next(d)
    data['Player_order'] = next(d)
    data['Player_Id'] = next(d)
    data['Difference'] = next(d)
    data['Difference_Efficiency_Per_Minutes'] = next(d)
    data['KD_Rate'] = next(d)
    data['Kills'] = next(d)
    data['Kills_Per_Minutes'] = next(d)
    data['Average_Open_Up_Efficiency'] = next(d)
    data['Average_Time_Spent_In_First_Kill'] = next(d)
    data['Ultimate_Kills'] = next(d)
    data['Ratio_Of_Ultimate_Kills_To_Total_Player_Kills'] = next(d)
    data['Team_Fight_Kills_Join_Rate'] = next(d)
    data['Kills_Per_Team_Fight'] = next(d)
    data['Deaths'] = next(d)
    data['Deaths_Per_Minutes'] = next(d)
    data['Live_Efficiency'] = next(d)
    data['Death_Team_Fight_Rate'] = next(d)
    data['Steals_Num'] = next(d)
    data['Fault_Num'] = next(d)
    data['Instant_Kills'] = next(d)
    data['Ultimate_Num'] = next(d)
    data['Ultimate_Using_Per_Minutes'] = next(d)
    data['Team_Fight_With_Ultimate_Ready_Rate'] = next(d)
    data['Ultimate_Succeed_Rate'] = next(d)
    data['Ultimate_Succeed_Rate_Offset'] = next(d)
    data['First_Blood_Get_Num'] = next(d)
    data['Num_Go_Back_Home_During_Ultimate'] = next(d)
    data['Deaths_During_Ultimate_Rate'] = next(d)
    data['Average_Time_Spent_In_Going_Back_Home'] = next(d)
    data['Self_Kills'] = next(d)
    data['Ultimate_Used_Times_In_Succeed_Team_Fight'] = next(d)
    data['Team_Fight_Failed_Num_Without_Ultimate'] = next(d)
    data['Team_Fight_Succeed_Num_Without_Ultimate'] = next(d)
    data['Num_Kill_Support'] = next(d)
    data['Num_Kill_Tank'] = next(d)
    data['Num_Kill_Fire'] = next(d)
    data['Num_Killed_By_Support'] = next(d)
    data['Num_Killed_By_Tank'] = next(d)
    data['Num_Killed_By_Fire'] = next(d)
    data['Num_Killed_By_Ultimate'] = next(d)
    data['Num_Kill_With_Head_Shot'] = next(d)
    data['Revive_Times'] = next(d)
    data['Tear_Down_Times'] = next(d)
    data['Be_Instant_Killed_Times'] = next(d)
    return data


def _is_valid_cell(sheet, row, col):
    if row < 0 or row >= sheet.nrows:
        return False
    if col < 0 or col >= sheet.ncols:
        return False
    return bool(cell_to_value(sheet, row, col))


def _next_valid_row_cell(sheet, row, col):
    if row is None:
        return None
    row = max(row, -1) + 1
    if row >= sheet.nrows:
        return None
    for i in range(row, sheet.nrows):
        if _is_valid_cell(sheet, i, col):
            return i
    return None


def _next_valid_col_cell(sheet, row, col):
    if col is None:
        return None
    col = max(col, -1) + 1
    if col >= sheet.ncols:
        return None
    for i in range(col, sheet.ncols):
        if _is_valid_cell(sheet, row, i):
            return i
    return None


def _get_next_col_data(sheet, row, col_begin, col_end):
    k = col_begin
    col_end = min(col_end, sheet.ncols)
    while k < col_end:
        yield cell_to_value(sheet, row, k)
        k = k + 1
    return None


def _get_next_row_data(sheet, col, row_begin, row_end):
    k = row_begin
    row_end = min(row_end, sheet.nrows)
    while k < row_end:
        yield cell_to_value(sheet, k, col)
        k = k + 1
    return None


def _read_single_player_global_statistics(sheet, row):
    data = dict()
    d = _get_next_col_data(sheet, row, 0, sheet.ncols)
    data['Play_Time'] = next(d)
    data['row_order'] = next(d)
    data['Player_Name'] = next(d)
    data['Player_Id'] = next(d)
    data['Difference'] = next(d)
    data['Difference_Efficiency_Per_10_Minutes'] = next(d)
    data['KD_Rate'] = next(d)
    data['Kills'] = next(d)
    data['Kills_Per_10_Minutes'] = next(d)
    data['Average_Open_Up_Efficiency'] = next(d)
    data['Ultimate_Kills'] = next(d)
    data['Ratio_Of_Ultimate_Kills_To_Total_Player_Kills'] = next(d)
    data['Kills_Per_Team_Fight'] = next(d)
    data['Deaths'] = next(d)
    data['Deaths_Per_10_Minutes'] = next(d)
    data['Death_Team_Fight_Rate'] = next(d)
    data['Average_Order_Of_Death'] = next(d)
    data['Steals_Num'] = next(d)
    data['Fault_Num'] = next(d)
    data['Instant_Kills'] = next(d)
    data['Ultimate_Num'] = next(d)
    data['Ultimate_Using_Per_10_Minutes'] = next(d)
    data['Ultimate_Succeed_Rate'] = next(d)
    data['Ultimate_Succeed_Rate_Offset'] = next(d)
    data['First_Blood_Get_Num'] = next(d)
    data['First_Blood_Give_Num_Per_10_Minutes'] = next(d)
    data['Fist_Kills_Get_Num_Per_10_Minutes'] = next(d)
    data['Die_In_Team_Fight_Rate'] = next(d)
    data['Num_Go_Back_Home_During_Ultimate'] = next(d)
    data['Deaths_During_Ultimate_Rate'] = next(d)
    data['Go_Back_Home_Num_Per_10_Minutes'] = next(d)
    data['Ultimate_Kills_Per_10_Minutes'] = next(d)
    data['Difference_Per_10_Minutes'] = next(d)
    data['Kills_Without_Ultimate'] = next(d)
    data['Death_Without_Ultimate'] = next(d)
    data['Difference_Without_Ultimate_Per_10_Minutes'] = next(d)
    data['Average_Time_Spent_To_Go_Back_Home_During_Ultimate'] = next(d)
    data['Self_Kills'] = next(d)
    data['Self_Kills_Per_10_Minutes'] = next(d)
    data['Team_Fight_Failed_Num_Without_Ultimate'] = next(d)
    data['Team_Fight_Succeed_Num_Without_Ultimate'] = next(d)
    data['Team_Fight_Succeed_Rate_Without_Ultimate'] = next(d)
    data['Ultimate_No_Ultimate_Efficiency_Difference'] = next(d)
    data['Team_Fight_Num_Joined'] = next(d)
    data['Num_Kill_Support'] = next(d)
    data['Num_Kill_Tank'] = next(d)
    data['Num_Kill_Fire'] = next(d)
    data['Num_Killed_By_Support'] = next(d)
    data['Num_Killed_By_Tank'] = next(d)
    data['Num_Killed_By_Fire'] = next(d)
    data['Num_Killed_By_Ultimate_'] = next(d)
    data['Num_Killed_With_Head_Shot'] = next(d)
    data['Head_Shot_Rate_Exclude_Instant_Kills'] = next(d)
    data['Direct_Damage_Ultimate_Using_Times'] = next(d)
    data['Direct_Damage_Ultimate_Killing_Times'] = next(d)
    data['Ultimate_Killing_Rate'] = next(d)
    data['Killed_By_Direct_Damage_Ultimate_Times'] = next(d)
    data['Direct_Damage_Ultimate_Dodge_Rate'] = next(d)
    data['Revive_Times'] = next(d)
    data['Ratio_Of_Revive_Times_To_Total_Revive_Times'] = next(d)
    data['Tear_Down_Times'] = next(d)
    data['Be_Instant_Killed_Times'] = next(d)
    return data


def _read_single_hero_global_statistics(sheet, row):
    data = dict()
    d = _get_next_col_data(sheet, row, 0, sheet.ncols)
    data['Hero_Id'] = next(d)
    data['Hero_Name'] = next(d)
    data['Kills'] = next(d)
    data['Ratio_Of_Kills_To_Total_Kills'] = next(d)
    data['First_Blood_Get_Num'] = next(d)
    data['Deaths'] = next(d)
    data['Ratio_Of_Deaths_To_Total_Deaths'] = next(d)
    data['First_Blood_Give_Num'] = next(d)
    data['Difference'] = next(d)
    data['Ultimate_Num'] = next(d)
    data['Ultimate_Kills'] = next(d)
    data['Ultimate_Efficiency'] = next(d)
    data['KD_Rate'] = next(d)
    data['First_Blood_Get_Rate'] = next(d)
    data['In_First_Team_Num'] = next(d)
    data['In_First_Team_Rate'] = next(d)
    data['In_Wait_List_Num'] = next(d)
    data['In_First_Team_Num_TeamA'] = next(d)
    data['In_First_Team_Num_TeamB'] = next(d)
    data['In_Wait_List_Num_TeamA'] = next(d)
    data['In_Wait_List_Num_TeamB'] = next(d)
    data['Playing_Time'] = next(d)
    data['Ratio_Of_Playing_Time_To_Whole_Time'] = next(d)
    data['Kills_Per_10_Minutes'] = next(d)
    data['Deaths_Per_10_Minutes'] = next(d)
    data['Ultimate_Using_Per_10_Minutes'] = next(d)
    data['First_Blood_Get_Num_Per_10_Minutes'] = next(d)
    data['First_Blood_Give_Num_Per_10_Minutes'] = next(d)
    data['Difference_Per_10_Minutes'] = next(d)
    data['FK_FD'] = next(d)
    data['Ultimate_Kills_Per_10_Minutes'] = next(d)
    data['Num_Go_Back_Home'] = next(d)
    data['Deaths_During_Ultimate_Rate'] = next(d)
    data['Go_Back_Home_Num_Per_10_Minutes'] = next(d)
    data['Ultimate_Win_Team_Fight_Num'] = next(d)
    data['Ultimate_Win_Team_Fight_Rate'] = next(d)
    data['Team_Fight_Failed_Num_Without_Ultimate'] = next(d)
    data['Team_Fight_Succeed_Num_Without_Ultimate'] = next(d)
    data['Team_Fight_Succeed_Rate_Without_Ultimate'] = next(d)
    data['Ultimate_No_Ultimate_Efficiency_Difference'] = next(d)
    return data


def cell_to_value(sheet, row, col):
    if not isinstance(sheet, xlrd.sheet.Sheet):
        return None

    cell = sheet.cell(row, col)
    if cell.ctype == xlrd.XL_CELL_ERROR:
        return None
    elif cell.ctype == xlrd.XL_CELL_DATE:
        d = xlrd.xldate_as_datetime(cell.value, 0)
        return d.hour * 3600 + d.minute * 60 + d.second if 0 <= cell.value < 1.0 else d
    else:
        return cell.value


def _read_single_battle_statistics(sheet, col):
    data = dict()
    data['Round'] = cell_to_value(sheet, 0, col)
    data['Map'] = cell_to_value(sheet, 1, col + 1)
    data['Winner'] = cell_to_value(sheet, 2, col + 1)
    data['Team_Fight_Num'] = cell_to_value(sheet, 3, col + 1)
    data['Average_Deaths_Per_Team_Fight'] = cell_to_value(sheet, 4, col + 1)
    data['First_Strike'] = cell_to_value(sheet, 5, col + 1)
    data['Total_Kills'] = cell_to_value(sheet, 6, col + 1)
    data['Total_Deaths'] = cell_to_value(sheet, 7, col + 1)
    data['Total_Time_Spent'] = cell_to_value(sheet, 8, col + 2)
    data['Average_Deaths_Per_Player'] = cell_to_value(sheet, 9, col + 1)
    data['Most_Deaths_Of_Player'] = cell_to_value(sheet, 10, col + 1)
    data['Most_Deaths_Player'] = cell_to_value(sheet, 10, col + 2)
    data['Least_Deaths_Of_Player'] = cell_to_value(sheet, 11, col + 1)
    data['Least_Deaths_Player'] = cell_to_value(sheet, 11, col + 2)
    data['Most_Last_Attack_Num_Of_Player'] = cell_to_value(sheet, 12, col + 1)
    data['Most_Last_Attack_Player'] = cell_to_value(sheet, 12, col + 2)
    data['Top_KD_Rate_Of_Player'] = cell_to_value(sheet, 13, col + 1)
    data['Top_KD_Rate_Player'] = cell_to_value(sheet, 13, col + 2)
    data['Average_Time_Spent_Per_Team_Fight'] = cell_to_value(sheet, 14, col + 1)
    data['Time_Spent_In_First_Team_Fight'] = cell_to_value(sheet, 15, col + 1)
    data['Total_Prepare_Time_Of_Team_Fight'] = cell_to_value(sheet, 16, col + 1)
    data['Average_Time_Gap_Between_Team_Fight'] = cell_to_value(sheet, 17, col + 1)
    data['Total_Time_Spent_In_Team_Fight'] = cell_to_value(sheet, 18, col + 1)
    data['Ratio_Of_Total_Team_Fight_Time_To_Total_Playing_Time'] = cell_to_value(sheet, 19, col + 1)
    data['Total_Ultimate_Num'] = cell_to_value(sheet, 20, col + 1)
    data['Average_Ultimate_Using_Per_Team_Fight'] = cell_to_value(sheet, 21, col + 1)
    data['Ultimate_Efficiency'] = cell_to_value(sheet, 22, col + 1)
    data['Top_Ultimate_Team_Fight_Victory_Efficiency_Of_Player'] = cell_to_value(sheet, 23, col + 1)
    data['Top_Ultimate_Team_Fight_Victory_Efficiency_Player'] = cell_to_value(sheet, 23, col + 2)
    data['Total_Ultimate_Kills'] = cell_to_value(sheet, 24, col + 1)
    data['Ratio_Of_Total_Ultimate_Kills_To_Total_Kills'] = cell_to_value(sheet, 25, col + 1)
    data['Most_Kills_Of_Hero'] = cell_to_value(sheet, 26, col + 1)
    data['Most_Kills_Hero'] = cell_to_value(sheet, 26, col + 2)
    data['Most_Deaths_Of_Hero'] = cell_to_value(sheet, 27, col + 1)
    data['Most_Deaths_Hero'] = cell_to_value(sheet, 27, col + 2)
    data['Top_Difference_Of_Hero'] = cell_to_value(sheet, 28, col + 1)
    data['Top_Difference_Hero'] = cell_to_value(sheet, 28, col + 2)
    data['Most_Ultimate_Using_Of_Hero'] = cell_to_value(sheet, 29, col + 1)
    data['Most_Ultimate_Using_Hero'] = cell_to_value(sheet, 29, col + 2)
    data['Most_Ultimate_Kills_Efficiency_Of_Hero'] = cell_to_value(sheet, 30, col + 1)
    data['Most_Ultimate_Kills_Efficiency_Hero'] = cell_to_value(sheet, 30, col + 2)
    data['TeamA_Total_Gathering_Time_After_Failure'] = cell_to_value(sheet, 31, col + 1)
    data['TeamA_Average_Gathering_Time_After_Failure'] = cell_to_value(sheet, 32, col + 1)
    data['TeamB_Total_Gathering_Time_After_Failure'] = cell_to_value(sheet, 33, col + 1)
    data['TeamB_Average_Gathering_Time_After_Failure'] = cell_to_value(sheet, 34, col + 1)
    data['Most_Fault_Times_Of_Player'] = cell_to_value(sheet, 35, col + 1)
    data['Most_Fault_Times_Player'] = cell_to_value(sheet, 35, col + 2)
    data['Most_Widen_Situation_Times_Of_Player'] = cell_to_value(sheet, 36, col + 1)
    data['Most_Widen_Situation_Times_Player'] = cell_to_value(sheet, 36, col + 2)
    data['TeamA_Team_Fights_Victory_Times'] = cell_to_value(sheet, 37, col + 1)
    data['TeamB_Team_Fights_Victory_Times'] = cell_to_value(sheet, 38, col + 1)
    data['TeamA_Direct_Damage_Ultimate_Using_Times'] = cell_to_value(sheet, 39, col + 1)
    data['TeamA_Direct_Damage_Ultimate_Missing_Times'] = cell_to_value(sheet, 40, col + 1)
    data['TeamB_Direct_Damage_Ultimate_Using_Times'] = cell_to_value(sheet, 41, col + 1)
    data['TeamB_Direct_Damage_Ultimate_Missing_Times'] = cell_to_value(sheet, 42, col + 1)
    return data


def _read_battle_summatize_statistics(sheet, col):
    data = dict()
    data['Round'] = cell_to_value(sheet, 0, col)
    data['Round_Num'] = cell_to_value(sheet, 1, col + 1)
    data['Winner'] = cell_to_value(sheet, 2, col + 1)
    data['Total_Team_Fight_Num'] = cell_to_value(sheet, 3, col + 1)
    data['Deaths_Per_10_Minutes'] = cell_to_value(sheet, 4, col + 1)
    data['Total_Self_Kills'] = cell_to_value(sheet, 5, col + 1)
    data['Total_Kills'] = cell_to_value(sheet, 6, col + 1)
    data['Total_Deaths'] = cell_to_value(sheet, 7, col + 1)
    data['Total_Time_Spent'] = cell_to_value(sheet, 8, col + 7)
    data['Average_Deaths_Per_Player'] = cell_to_value(sheet, 9, col + 1)
    data['Most_Last_Attack_Num_Of_Player'] = cell_to_value(sheet, 12, col + 1)
    data['Most_Last_Attack_Player'] = cell_to_value(sheet, 12, col + 2)
    data['Average_Time_Spent_Per_Team_Fight'] = cell_to_value(sheet, 14, col + 1)
    data['Average_Time_Spent_In_First_Team_Fight'] = cell_to_value(sheet, 15, col + 1)
    data['Average_Prepare_Time_Of_Team_Fight'] = cell_to_value(sheet, 16, col + 1)
    data['Average_Time_Gap_Between_Team_Fight'] = cell_to_value(sheet, 17, col + 1)
    data['Average_Time_Spent_In_All_Team_Fight'] = cell_to_value(sheet, 18, col + 1)
    data['Ratio_Of_Total_Team_Fight_Time_To_Total_Playing_Time'] = cell_to_value(sheet, 19, col + 1)
    data['Total_Ultimate_Num'] = cell_to_value(sheet, 20, col + 1)
    data['Average_Ultimate_Using_Per_Team_Fight'] = cell_to_value(sheet, 21, col + 1)
    data['Team_Fight_Win_Rate_Of_First_Blood_Owner_Team'] = cell_to_value(sheet, 22, col + 1)
    data['Top_KD_Rate_Of_Hero'] = cell_to_value(sheet, 23, col + 1)
    data['Top_KD_Rate_Hero'] = cell_to_value(sheet, 23, col + 2)
    data['Total_Ultimate_Kills'] = cell_to_value(sheet, 24, col + 1)
    data['Ratio_Of_Total_Ultimate_Kills_To_Total_Kills'] = cell_to_value(sheet, 25, col + 1)
    data['Most_Kills_Of_Hero'] = cell_to_value(sheet, 26, col + 1)
    data['Most_Kills_Hero'] = cell_to_value(sheet, 26, col + 2)
    data['Most_Deaths_Of_Hero'] = cell_to_value(sheet, 27, col + 1)
    data['Most_Deaths_Hero'] = cell_to_value(sheet, 27, col + 2)
    data['Top_Difference_Of_Hero'] = cell_to_value(sheet, 28, col + 1)
    data['Top_Difference_Hero'] = cell_to_value(sheet, 28, col + 2)
    data['Most_Ultimate_Using_Of_Hero'] = cell_to_value(sheet, 29, col + 1)
    data['Most_Ultimate_Using_Hero'] = cell_to_value(sheet, 29, col + 2)
    data['Most_Ultimate_Kills_Efficiency_Of_Hero'] = cell_to_value(sheet, 30, col + 1)
    data['Most_Ultimate_Kills_Efficiency_Hero'] = cell_to_value(sheet, 30, col + 2)
    data['TeamA_Average_Gathering_Time_After_Failure'] = cell_to_value(sheet, 31, col + 1)
    data['TeamB_Average_Gathering_Time_After_Failure'] = cell_to_value(sheet, 32, col + 1)
    data['Most_Fault_Times_Of_Player'] = cell_to_value(sheet, 33, col + 1)
    data['Most_Fault_Times_Player'] = cell_to_value(sheet, 33, col + 2)
    data['Most_First_Blood_Num_Get_By_Single_Player'] = cell_to_value(sheet, 34, col + 1)
    data['Most_First_Blood_Num_Get_Player'] = cell_to_value(sheet, 34, col + 2)
    data['Top_KD_Rate_Of_Player'] = cell_to_value(sheet, 35, col + 1)
    data['Top_KD_Rate_Player'] = cell_to_value(sheet, 35, col + 2)
    data['Top_Difference_Of_Player'] = cell_to_value(sheet, 36, col + 1)
    data['Top_Difference_Player'] = cell_to_value(sheet, 36, col + 2)
    data['TeamA_Team_Fights_Victory_Times'] = cell_to_value(sheet, 37, col + 1)
    data['TeamA_Team_Fights_Victory_Rate'] = cell_to_value(sheet, 37, col + 3)
    data['TeamB_Team_Fights_Victory_Times'] = cell_to_value(sheet, 38, col + 1)
    data['TeamB_Team_Fights_Victory_Rate'] = cell_to_value(sheet, 38, col + 3)
    data['TeamA_Name'] = cell_to_value(sheet, 39, col + 1)
    data['TeamA_Direct_Damage_Ultimate_Using_Times'] = cell_to_value(sheet, 39, col + 5)
    data['TeamA_Direct_Damage_Ultimate_Missing_Times'] = cell_to_value(sheet, 39, col + 7)
    data['TeamA_Direct_Damage_Ultimate_Hit_Rate'] = cell_to_value(sheet, 39, col + 9)
    data['TeamA_Direct_Damage_Ultimate_Dodge_Rate'] = cell_to_value(sheet, 39, col + 11)
    data['TeamB_Name'] = cell_to_value(sheet, 40, col + 1)
    data['TeamB_Direct_Damage_Ultimate_Using_Times'] = cell_to_value(sheet, 40, col + 5)
    data['TeamB_Direct_Damage_Ultimate_Missing_Times'] = cell_to_value(sheet, 40, col + 7)
    data['TeamB_Direct_Damage_Ultimate_Hit_Rate'] = cell_to_value(sheet, 40, col + 9)
    data['TeamB_Direct_Damage_Ultimate_Dodge_Rate'] = cell_to_value(sheet, 40, col + 11)
    return data


def _read_battle_statistics(sheet, col_begin):
    if cell_to_value(sheet, 0, col_begin) == 'ALL':
        return _read_battle_summatize_statistics(sheet, col_begin)
    return _read_single_battle_statistics(sheet, col_begin)
