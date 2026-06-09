import datetime
from dateutil import relativedelta
from tkcalendar import Calendar
from babel import numbers
import os
import sys
import csv
import glob
import pandas as pd
import mysql.connector
import openpyxl
from tkinter import *
import tkinter.messagebox
from pathlib import Path
cur_dir = Path(__file__).parent
sys.path.append(str(cur_dir / '../../../../../../Credentials'))
import credentials

# ===============================================================================
# Connect to Db
# ===============================================================================


def connect_read():
    conn = mysql.connector.connect(host=credentials.host_read,
                                   database=credentials.database_hq,
                                   user=credentials.username,
                                   password=credentials.password)

    return conn

# ===============================================================================
# Reg Payperiod Hours Query
# ===============================================================================

def get_payperiod_hours(conn_read, st_date):
    query = f'''
SELECT olocation_name              AS Loc,
       ouser_department_name       AS Dept,
       ouser_last_name             AS Last,
       ouser_first_name            AS First,
       DATE(pay_period_start_date) AS StartDate,
       DATE(pay_period_end_date)   AS EndDate,
       time_card_is_approved       AS Approved,
       time_card_manager_approved  AS ManagerApproved,
       sum(wr.RegHours)                 AS Regular,
       GREATEST(0, sum(wr.OTHours))     AS OT,
       sum(wr.DoubleHours)             AS DoubleOT,
       SUM(PTOHours)               AS PTO,
       SUM(CharitibleTO)           AS CTO,
       SUM(UnpaidTO)               AS UPTO,
       SUM(MaternityPTO)           AS Maternity,
       SUM(PaternityPTO)           AS Paternity,
       SUM(Sickleave)              AS SickLeave,
       SUM(`COVID Sickleave`)      AS 'COVID Sickleave',
       SUM(BereavementPTO)         AS Bereavement,
       SUM(TravelHours)            AS Travel,
       SUM(HolidayHours)           AS Holiday,
       SUM(FMLA)                   AS FMLA,
       SUM(HealthTimeOff)          AS 'Personal Health',
       SUM(FloatingHoliday)        AS 'Floating Holiday',
       sum(wr.TotHours)                 AS Total
FROM (SELECT tc.ouser_id,
             olocation_name,
             IFNULL(ud2.ouser_department_name, ud.ouser_department_name)                            AS ouser_department_name,
             ud2.ouser_department_id,
             ud2.ouser_department_sub_of,
             udl.ouser_department_id                                                                AS TopDepartment,
             ouser_last_name,
             ouser_first_name,
             pay_period_start_date,
             pay_period_end_date,
             CASE WHEN tp.time_punch_type_id = 2 THEN 1 ELSE tp.time_punch_type_id END              AS time_punch_type_id,
             time_card_is_approved,
             time_card_manager_approved,
             ROUND(SUM(CASE WHEN time_punch_type_id IN (1, 2) THEN time_punch_hours ELSE 0 END), 2) AS StandardHours,
             ROUND(SUM(CASE WHEN time_punch_type_id = 3 THEN time_punch_hours ELSE 0 END), 2)       AS PTOHours,
             ROUND(SUM(CASE WHEN time_punch_type_id = 4 THEN time_punch_hours ELSE 0 END), 2)       AS UnpaidTO,
             ROUND(SUM(CASE WHEN time_punch_type_id = 10 THEN time_punch_hours ELSE 0 END), 2)      AS CharitibleTO,
             ROUND(SUM(CASE WHEN time_punch_type_id = 12 THEN time_punch_hours ELSE 0 END), 2)      AS Sickleave,
             ROUND(SUM(CASE WHEN time_punch_type_id = 13 THEN time_punch_hours ELSE 0 END), 2)      AS 'COVID Sickleave',
             ROUND(SUM(CASE WHEN time_punch_type_id = 7 THEN time_punch_hours ELSE 0 END), 2)       AS MaternityPTO,
             ROUND(SUM(CASE WHEN time_punch_type_id = 8 THEN time_punch_hours ELSE 0 END), 2)       AS PaternityPTO,
             ROUND(SUM(CASE WHEN time_punch_type_id = 9 THEN time_punch_hours ELSE 0 END), 2)       AS BereavementPTO,
             ROUND(SUM(CASE WHEN time_punch_type_id = 5 THEN time_punch_hours ELSE 0 END), 2)       AS TravelHours,
             ROUND(SUM(CASE WHEN time_punch_type_id = 6 THEN time_punch_hours ELSE 0 END), 2)       AS HolidayHours,
             ROUND(SUM(CASE WHEN time_punch_type_id = 11 THEN time_punch_hours ELSE 0 END), 2)      AS FMLA,
             ROUND(SUM(CASE WHEN time_punch_type_id = 16 THEN time_punch_hours ELSE 0 END), 2)      AS HealthTimeOff,
             ROUND(SUM(CASE WHEN time_punch_type_id = 17 THEN time_punch_hours ELSE 0 END), 2)      AS FloatingHoliday,
             ROUND(SUM(time_punch_hours), 2)
      FROM hq.oTimePunches tp
                JOIN hq.oTimeCards tc ON tc.time_card_id = tp.time_card_id
                JOIN hq.oPayPeriods pp ON pp.pay_period_id = tc.pay_period_id
                JOIN hq.ousers u ON u.ouser_id = tc.ouser_id
                LEFT JOIN hq.oUserDepartments ud ON u.ouser_department_id = ud.ouser_department_id
                LEFT JOIN hq.oUserDepartments ud2 ON ud.ouser_department_sub_of = ud2.ouser_department_id
                LEFT JOIN hq.oUserDepartmentLinks udl ON u.ouser_id = udl.ouser_id AND udl.ouser_department_level = 1
                LEFT JOIN hq.oTimeOffRequestDays tord ON tord.time_punch_id = tp.time_punch_id
                LEFT JOIN hq.oTimeOffRequests tor ON tor.time_off_request_id = tord.time_off_request_id
                JOIN hq.olocations l ON l.olocation_id = u.ouser_location
                LEFT JOIN (SELECT u.ouser_id, utd.ouser_management_tier_description
                           FROM hq.ousers u
                                     LEFT JOIN hq.oUserTiers ut ON u.ouser_id = ut.ouser_id
                                     LEFT JOIN hq.oUserTierDescriptions utd
                                               ON u.ouser_department_id = utd.ouser_department_id AND
                                                  ut.ouser_management_tier_id = utd.ouser_management_tier_id) up
                          ON up.ouser_id = u.ouser_id
      WHERE time_punch_is_void = 0
        AND pay_period_start_date = '{st_date}'
        AND time_punch_type_id != 4
        AND ouser_is_hourly = 1
      GROUP BY tc.ouser_id, ouser_last_name, ouser_first_name, time_card_is_approved,
               time_card_manager_approved, CASE WHEN tp.time_punch_type_id = 2 THEN 1 ELSE tp.time_punch_type_id END) ht
          LEFT JOIN (SELECT tc.ouser_id,
                            ROUND(SUM(time_card_rollup_regular_hours), 2)            AS RegHours,
                            ROUND(SUM(time_card_rollup_overtime_hours), 2)           AS OTHours,
                            ROUND(SUM(time_card_rollup_doubletime_hours), 2)         AS DoubleHours,
                            ROUND(SUM(time_card_rollup_total_hours), 2)              AS TotHours,
                            ROUND(SUM(time_card_rollup_weekend_regular_hours), 2)    AS WeekEndReg,
                            ROUND(SUM(time_card_rollup_weekend_overtime_hours), 2)   AS WeekEndOT,
                            ROUND(SUM(time_card_rollup_weekend_doubletime_hours), 2) AS WeekEndDOT
                     FROM hq.oTimeCardWeekRollups tcwr
                               JOIN hq.oTimeCards tc ON tcwr.time_card_id = tc.time_card_id
                               JOIN hq.oPayPeriods pp ON pp.pay_period_id = tc.pay_period_id
                               JOIN hq.ousers u ON u.ouser_id = tc.ouser_id
                     WHERE pay_period_start_date = '{st_date}'
                       AND ouser_is_hourly = 1
                     GROUP BY tc.ouser_id) wr ON wr.ouser_id = ht.ouser_id AND time_punch_type_id = 1
GROUP BY olocation_name, ouser_department_sub_of, ouser_last_name, ouser_first_name;
            '''
    pp_hours_df = pd.read_sql(query, conn_read)

    return pp_hours_df

# ===============================================================================
# Kronos upload
# ===============================================================================

def get_payperiod_kronos(conn_read, st_date):
    query = f'''
SELECT ouser_employee_number                          AS 'Employee ID',
       CONCAT(ouser_first_name, ' ', ouser_last_name) AS 'Employee Name',
       DATE(pay_period_start_date)                    AS 'Pay Period Start Date',
       DATE(pay_period_end_date)                      AS 'Pay Period End Date',
       'Regular'                                      AS 'Pay Statement Type',
       ''                                             AS 'Pay Statement Type Key',
       ''                                             AS 'Pay Order',
       sum(wr.RegHours)                                    AS 'E Reg Hours',
       GREATEST(0, sum(wr.OTHours))                        AS 'E OT Hours',
       sum(wr.DoubleHours)                                 AS 'E Double OT Hours',
       SUM(PTOHours)                                  AS 'E PTO Hours',
       SUM(CharitibleTO)                              AS 'E CTO Hours',
       SUM(MaternityPTO)                              AS 'E Maternity Hours',
       SUM(PaternityPTO)                              AS 'E Paternity Hours',
       SUM(BereavementPTO)                            AS 'E Bereavement Hours',
       SUM(TravelHours)                               AS 'E Travel Hours',
       SUM(HolidayHours)                              AS 'E Holiday Hours',
       SUM(SickPTO)                                   AS 'E Sick Leave Hours',
       SUM(`COVID Sickleave`)                         AS 'E COVID Sick Leave Hours',
       SUM(HealthTimeOff)                             AS 'E Personal Health',
       SUM(FloatingHoliday)                           AS 'E Floating Holiday'
FROM (SELECT tc.ouser_id,
             olocation_name,
             ud2.ouser_department_name,
             ud2.ouser_department_id,
             ud2.ouser_department_sub_of,
             udl.ouser_department_id                                                                AS TopDepartment,
             ouser_last_name,
             ouser_first_name,
             ouser_employee_number,
             pay_period_start_date,
             pay_period_end_date,
             CASE WHEN tp.time_punch_type_id = 2 THEN 1 ELSE tp.time_punch_type_id END              AS time_punch_type_id,
             time_card_is_approved,
             time_card_manager_approved,
             ROUND(SUM(CASE WHEN time_punch_type_id IN (1, 2) THEN time_punch_hours ELSE 0 END), 2) AS StandardHours,
             ROUND(SUM(CASE WHEN time_punch_type_id = 3 THEN time_punch_hours ELSE 0 END), 2)       AS PTOHours,
             ROUND(SUM(CASE WHEN time_punch_type_id = 4 THEN time_punch_hours ELSE 0 END), 2)       AS UnpaidTO,
             ROUND(SUM(CASE WHEN time_punch_type_id = 10 THEN time_punch_hours ELSE 0 END), 2)      AS CharitibleTO,
             ROUND(SUM(CASE WHEN time_punch_type_id = 12 THEN time_punch_hours ELSE 0 END), 2)      AS SickPTO,
             ROUND(SUM(CASE WHEN time_punch_type_id = 13 THEN time_punch_hours ELSE 0 END),
                   2)                                                                               AS 'COVID Sickleave',
             ROUND(SUM(CASE WHEN time_punch_type_id = 7 THEN time_punch_hours ELSE 0 END), 2)       AS MaternityPTO,
             ROUND(SUM(CASE WHEN time_punch_type_id = 8 THEN time_punch_hours ELSE 0 END), 2)       AS PaternityPTO,
             ROUND(SUM(CASE WHEN time_punch_type_id = 9 THEN time_punch_hours ELSE 0 END), 2)       AS BereavementPTO,
             ROUND(SUM(CASE WHEN time_punch_type_id = 5 THEN time_punch_hours ELSE 0 END), 2)       AS TravelHours,
             ROUND(SUM(CASE WHEN time_punch_type_id = 6 THEN time_punch_hours ELSE 0 END), 2)       AS HolidayHours,
             ROUND(SUM(CASE WHEN time_punch_type_id = 16 THEN time_punch_hours ELSE 0 END), 2)      AS HealthTimeOff,
             ROUND(SUM(CASE WHEN time_punch_type_id = 17 THEN time_punch_hours ELSE 0 END), 2)      AS FloatingHoliday,
             ROUND(SUM(time_punch_hours), 2)
      FROM hq.oTimePunches tp
                JOIN hq.oTimeCards tc ON tc.time_card_id = tp.time_card_id
                JOIN hq.oPayPeriods pp ON pp.pay_period_id = tc.pay_period_id
                JOIN hq.ousers u ON u.ouser_id = tc.ouser_id
                LEFT JOIN hq.oUserDepartments ud ON u.ouser_department_id = ud.ouser_department_id
                LEFT JOIN hq.oUserDepartments ud2 ON ud.ouser_department_sub_of = ud2.ouser_department_id
                LEFT JOIN hq.oUserDepartmentLinks udl ON u.ouser_id = udl.ouser_id AND udl.ouser_department_level = 1
                LEFT JOIN hq.oTimeOffRequestDays tord ON tord.time_punch_id = tp.time_punch_id
                LEFT JOIN hq.oTimeOffRequests tor ON tor.time_off_request_id = tord.time_off_request_id
                JOIN hq.olocations l ON l.olocation_id = u.ouser_location
      WHERE time_punch_is_void = 0
        AND pay_period_start_date = '{st_date}'
        AND time_punch_type_id != 4
        AND ouser_is_hourly = 1
      GROUP BY tc.ouser_id, ouser_last_name, ouser_first_name, time_card_is_approved,
               time_card_manager_approved, CASE WHEN tp.time_punch_type_id = 2 THEN 1 ELSE tp.time_punch_type_id END) ht
          LEFT JOIN (SELECT tc.ouser_id,
                            ROUND(SUM(time_card_rollup_regular_hours), 2)            AS RegHours,
                            ROUND(SUM(time_card_rollup_overtime_hours), 2)           AS OTHours,
                            ROUND(SUM(time_card_rollup_doubletime_hours), 2)         AS DoubleHours,
                            ROUND(SUM(time_card_rollup_total_hours), 2)              AS TotHours,
                            ROUND(SUM(time_card_rollup_weekend_regular_hours), 2)    AS WeekendReg,
                            ROUND(SUM(time_card_rollup_weekend_overtime_hours), 2)   AS WeekendOT,
                            ROUND(Sum(time_card_rollup_weekend_doubletime_hours), 2) As WeekendDOT
                     FROM hq.oTimeCardWeekRollups tcwr
                               JOIN hq.oTimeCards tc ON tcwr.time_card_id = tc.time_card_id
                               JOIN hq.oPayPeriods pp ON pp.pay_period_id = tc.pay_period_id
                               JOIN hq.ousers u ON u.ouser_id = tc.ouser_id
                     WHERE pay_period_start_date = '{st_date}'
                       AND ouser_is_hourly = 1
                     GROUP BY tc.ouser_id) wr ON wr.ouser_id = ht.ouser_id AND time_punch_type_id = 1
GROUP BY olocation_name, ouser_department_sub_of, ouser_last_name, ouser_first_name;
            '''
    pp_kronos_df = pd.read_sql(query, conn_read)

    return pp_kronos_df

# ===============================================================================
# Hours Per Week - 1 Month
# ===============================================================================
def get_hours_1_month(conn_read):
    query = f'''
                SELECT
                    u.ouser_id,
                    IFNULL(ud2.ouser_department_name, ud.ouser_department_name) AS Dept,
                    ouser_last_name AS 'Last Name',
                    ouser_first_name AS 'First Name',
                    l.olocation_name AS Location,
                    MIN(pay_period_start_date) AS FirstPayDate,
                    MAX(pay_period_end_date) AS LastPayDate,
                    ROUND(SUM(time_card_rollup_regular_hours),2) AS RegularHours,
                    ROUND(SUM(time_card_rollup_overtime_hours),2) AS OvertimeHours,
                    ROUND(SUM((time_card_rollup_regular_hours)+(time_card_rollup_other_hours)+(time_card_rollup_overtime_hours)),2) AS TotalHours
                FROM hq.oTimeCardWeekRollups tcwr
                        JOIN hq.oTimeCards tc ON tcwr.time_card_id = tc.time_card_id
                        JOIN hq.oPayPeriods pp ON pp.pay_period_id = tc.pay_period_id
                        JOIN hq.ousers u ON u.ouser_id = tc.ouser_id
                        LEFT JOIN oUserDepartments ud ON u.ouser_department_id = ud.ouser_department_id
                        LEFT JOIN oUserDepartments ud2 ON ud.ouser_department_sub_of = ud2.ouser_department_id
                        JOIN hq.olocations l ON u.ouser_location = l.olocation_id
                WHERE pay_period_start_date >= CURDATE() - INTERVAL 6 WEEK
                AND pay_period_end_date < CURDATE()
                AND pay_period_start_date >= '2018-01-01'
                AND ouser_is_hourly = 1
                AND ouser_status = 'active'
                GROUP BY
                    Dept,
                    ouser_last_name,
                    ouser_first_name;
            '''
    pp_1month_df = pd.read_sql(query, conn_read)

    return pp_1month_df

# ===============================================================================
# Hours Per Week - 6 Month
# ===============================================================================
def get_hours_6_month(conn_read):
    query = f'''
                SELECT
                    u.ouser_id,
                    IFNULL(ud2.ouser_department_name, ud.ouser_department_name) AS Dept,
                    ouser_last_name AS 'Last Name',
                    ouser_first_name AS 'First Name',
                    l.olocation_name AS Location,
                    MIN(pay_period_start_date) AS FirstPayDate,
                    MAX(pay_period_end_date) AS LastPayDate,
                    ROUND(SUM(time_card_rollup_regular_hours),2) AS RegularHours,
                    ROUND(SUM(time_card_rollup_overtime_hours),2) AS OvertimeHours,
                    ROUND(SUM((time_card_rollup_regular_hours)+(time_card_rollup_other_hours)+(time_card_rollup_overtime_hours)),2) AS TotalHours
                FROM hq.oTimeCardWeekRollups tcwr
                        JOIN hq.oTimeCards tc ON tcwr.time_card_id = tc.time_card_id
                        JOIN hq.oPayPeriods pp ON pp.pay_period_id = tc.pay_period_id
                        JOIN hq.ousers u ON u.ouser_id = tc.ouser_id
                        LEFT JOIN oUserDepartments ud ON u.ouser_department_id = ud.ouser_department_id
                        LEFT JOIN oUserDepartments ud2 ON ud.ouser_department_sub_of = ud2.ouser_department_id
                        JOIN hq.olocations l ON u.ouser_location = l.olocation_id
                WHERE pay_period_start_date >= CURDATE() - INTERVAL 6 MONTH
                AND pay_period_end_date < CURDATE()
                AND pay_period_start_date >= '2018-01-01'
                AND ouser_is_hourly = 1
                AND ouser_status = 'active'
                GROUP BY
                    Dept,
                    ouser_last_name,
                    ouser_first_name;
            '''
    pp_6months_df = pd.read_sql(query, conn_read)

    return pp_6months_df

# ===============================================================================
# Hours by Week (Hazard)
# ===============================================================================

def get_hours_by_week(conn_read):
    query = f'''
                SELECT
                    l.olocation_name AS Location,
                    ug.ouser_group_name AS Dept,
                #      tc.time_card_id,
                    u.ouser_first_name AS First,
                    u.ouser_last_name AS Last,
                    dt.adate - INTERVAL 6 DAY AS StartDate,
                    dt.adate AS EndDate,

                #      tcwr.time_card_rollup_year AS Years,
                    ROUND(SUM(tcwr.time_card_rollup_regular_hours),2) AS Regular,
                #      ROUND(SUM(time_card_rollup_overtime_hours),2) AS OvertimeHours,
                    ROUND(SUM(tcwr.time_card_rollup_overtime_hours),2) AS Overtime,
                    ROUND(SUM(tcwr.time_card_rollup_doubletime_hours),2)AS 'Double OT',
                    ROUND(SUM((tcwr.time_card_rollup_regular_hours)+(tcwr.time_card_rollup_overtime_hours)+ (tcwr.time_card_rollup_doubletime_hours)),2) AS TotalHours
                FROM hq.oTimeCardWeekRollups tcwr
                        JOIN hq.oTimeCards tc ON tcwr.time_card_id = tc.time_card_id
                        JOIN hq.ousers u ON u.ouser_id = tc.ouser_id
                        JOIN hq.ousergroups ug ON ug.ouser_group_id = u.ouser_group
                        JOIN hq.oUserSubGroups usg ON usg.ouser_sub_group_id = u.ouser_sub_group_id
                        JOIN hq.olocations l ON u.ouser_location = l.olocation_id
                        JOIN (SELECT MAX(adate) AS adate, WEEK(adate) week_num, YEAR(adate) year_num,
                                    CASE WHEN WEEK(adate) = WEEK(CURDATE()) THEN 1
                                            WHEN WEEK(adate) = WEEK(CURDATE())-1 THEN 2
                                            WHEN WEEK(adate) = WEEK(CURDATE())-2 THEN 3
                                            WHEN WEEK(adate) = WEEK(CURDATE())-3 THEN 4
                                            WHEN WEEK(adate) = WEEK(CURDATE())-4 THEN 5 END AS Week_Order
                                FROM analytics.adates
                                WHERE adate <= CURDATE()
                                AND adate>= LAST_DAY(CURDATE() - INTERVAL 1 MONTH) - INTERVAL 3 WEEK
                                GROUP BY WEEK(adate))dt ON dt.week_num = tcwr.time_card_rollup_week
                    AND dt.year_num = tcwr.time_card_rollup_year
                WHERE 1=1
                AND ouser_is_hourly = 1
                #   AND u.ouser_group IN (8,17,18)
                AND dt.adate >= LAST_DAY(CURDATE() - INTERVAL 1 MONTH) - INTERVAL 3 WEEK
                AND dt.Week_Order <= 3
                #   AND dt.week_num <> WEEK(CURDATE())
                #   AND ouser_status = 'active'
                GROUP BY tcwr.time_card_rollup_week, u.ouser_id;
            '''
    hours_wk_df = pd.read_sql(query, conn_read)

    return hours_wk_df

# ===============================================================================
# Health & Floating Holiday
# ===============================================================================
def get_health_holiday(conn_read, st_date):
    query = f'''
                SELECT u.ouser_first_name AS 'First Name', u.ouser_last_name AS 'Last Name', 
             DATE(tp.time_punch_in) AS 'Week Of',
             IF(tp.time_punch_type_id = 16, SUM(tp.time_punch_hours), 0) AS 'Personal Health',
             IF(tp.time_punch_type_id = 17, SUM(tp.time_punch_hours), 0) AS 'Floating Holiday',
             l.olocation_name AS Location
        FROM hq.oTimePunches tp
          JOIN hq.oTimeCards tc ON tc.time_card_id = tp.time_card_id
          JOIN hq.ousers u ON tc.ouser_id = u.ouser_id
          LEFT JOIN hq.olocations l ON u.ouser_location = l.olocation_id
        WHERE True
            AND tp.time_punch_in >= DATE('{st_date}' - INTERVAL WEEKDAY('{st_date}') DAY)
            AND tp.time_punch_in <= DATE('{st_date}' - INTERVAL WEEKDAY('{st_date}') DAY) + INTERVAL 6 DAY
            AND tp.time_punch_type_id IN (16, 17)
            AND tp.time_punch_is_void = 0
        GROUP BY 'Week Of', u.ouser_id;
            '''
    pp_health_holiday_df = pd.read_sql(query, conn_read)

    return pp_health_holiday_df


# ===============================================================================
# Date Select 
# ===============================================================================

def filter_date(event, cal):
    index_values = cal.selection_get() 
    print(index_values)
    start_date = datetime.datetime.strftime(index_values, '%Y-%m-%d')
    date_inputs.config(text= f"{start_date}")
    print(start_date)

    if index_values >= (datetime.datetime.today()).date():
        txt = "Future date selected \n" \
              "This will not return any data"
        pop_up(txt)

# ===============================================================================
# Completion PopUp
# ===============================================================================

def pop_up(text):
    tkinter.messagebox.showinfo("pop_up", text)

# ===============================================================================
# Export to Excel
# ===============================================================================


def export(original_path):
    date_select = date_inputs.cget("text")
    original_path, payperiod_hours, kronos, hours_1_month, hours_6_month, hours_by_week, health_holiday = main(date_select)
    os.chdir(original_path)
    downloads_path = str(Path.home() / "Downloads")
    os.chdir(downloads_path)

    file_name = 'Pay Period Hours by User.xlsx'

    sheet = [{'Name': 'Pay period', 'df': payperiod_hours, 'max_row': payperiod_hours.shape[0],
    'max_col': (payperiod_hours.shape[1] - 1)},
    {'Name': 'Hours per Week - 1 month', 'df': hours_1_month, 'max_row': hours_1_month.shape[0],
    'max_col': (hours_1_month.shape[1] - 1)},
    {'Name': 'Hours per Week - 6 months', 'df': hours_6_month, 'max_row': hours_6_month.shape[0],
    'max_col': (hours_6_month.shape[1] - 1)},
    {'Name': 'Hours by Week', 'df': hours_by_week, 'max_row': hours_by_week.shape[0],
    'max_col': (hours_by_week.shape[1] - 1)},
    {'Name': 'Pay Period Kronos', 'df': kronos, 'max_row': kronos.shape[0],
    'max_col': (kronos.shape[1] - 1)},
    {'Name': 'Health & Floating Holiday', 'df': health_holiday, 'max_row': health_holiday.shape[0],
    'max_col': (health_holiday.shape[1] - 1)}
    ]

    writer = pd.ExcelWriter(file_name, engine='xlsxwriter')

    # iteration through sheets
    for worksheets in sheet:
        worksheets['df'].to_excel(writer, sheet_name=worksheets['Name'], startrow=1, header=False, index=False)
        worksheet = writer.sheets[worksheets['Name']]
        column_settings = []

    # getting list of headers
        for header in (worksheets['df']).columns:
            column_settings.append({'header': header})

        # setting table conditions
        worksheet.add_table(0, 0, worksheets['max_row'], worksheets['max_col'],
        {'columns': column_settings,
        'style': 'Table Style Medium 2'})

        # setting column widths
        for column in worksheets['df']:
            column_width = 20
            col_idx = (worksheets['df']).columns.get_loc(column)
            worksheet.set_column(col_idx, col_idx, column_width)

    writer.close()
    text = 'Please check your downloads \n ' \
           'for: Pay Period Hours by User.xlsx'
    pop_up(text)

# ===============================================================================
# Main
# ===============================================================================


def main(date_select):
    original_path = os.getcwd()
    print('removing file from downloads')
    conn_read = connect_read()
    print('getting reads from db')
    payperiod_hours = get_payperiod_hours(conn_read, date_select)
    kronos = get_payperiod_kronos(conn_read, date_select)
    hours_1_month = get_hours_1_month(conn_read)
    hours_6_month = get_hours_6_month(conn_read)
    hours_by_week = get_hours_by_week(conn_read)
    health_holiday = get_health_holiday(conn_read, date_select)
    conn_read.close()

    return original_path, payperiod_hours, kronos, hours_1_month, hours_6_month, hours_by_week, health_holiday

# ===============================================================================
# INIT
# ===============================================================================


if __name__ == '__main__':
    root = Tk()
    root.title('PayPeriod Hours by User.xlsx')
    root.iconbitmap({filepath_to_icon})
    # Set Geometry
    root.geometry("350x350")

    original_path = str(Path.home() / "Downloads")

    # Add Calendar
    cal = Calendar(root,date_patternstr = 'y-mm-dd', selectmode = 'day')
    cal.place(relx=.15, rely=.25)

    cal.bind("<<CalendarSelected>>", lambda event: filter_date(event, cal))

    #print("selected date: ", st_date)

    intro_txt = Label(root, text='Select Start Date:', font=("Arial", 14))
    intro_txt.place(relx=.30, rely=.15)

    date_inputs = Label(root, text='', font=("Arial", 8))
    date_inputs.place(relx=.70, rely=.80)

    Retrieve_button = Button(root, text="Retrieve File",
                             font=("Arial", 10), command=lambda: export(original_path))
    Retrieve_button.place(relx=.40, rely=.80)

    root.mainloop()
