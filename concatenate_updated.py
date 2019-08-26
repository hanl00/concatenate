import os
import glob
import pandas as pd
import timeit

start = timeit.default_timer()
current_path = os.getcwd()
print(current_path)
extension = 'xlsx'
final_course_file = glob.glob('*.{}'.format(extension))[0]
final_institution_file = glob.glob('*.{}'.format(extension))[1]

final_course_file_path = os.path.abspath(final_course_file)
final_institution_file_path = os.path.abspath(final_institution_file)

# open the final_file.xlsx
final_course_data = pd.read_excel(final_course_file_path)
final_institution_data = pd.read_excel(final_institution_file_path)

# search for all excel files in the 'excel files' folder
excel_files_path = os.getcwd() + '\excel files'
extension = 'xlsm'
os.chdir(excel_files_path)
fileList = glob.glob('*.{}'.format(extension))


for fileName in fileList:
    # obtaining file directory
    fileDir = os.path.realpath(fileName)

    # load both excel files
    DataFrame = pd.read_excel(fileDir)
    DataFrame_institution = pd.read_excel(fileDir, sheet_name='Institution Details')

    # Institution
    DataFrame_institution.rename(columns={'INSTITUTION DETAILS': 'Univeristy_Name', 'Unnamed: 2': 'Country',
                                          'Unnamed: 3': 'Int_Ph_num', 'Unnamed: 4': 'Int_Emails',
                                          'Unnamed: 5': 'Website', 'Unnamed: 6': 'T_num_of_stu',
                                          'Unnamed: 7': 'Address', 'Unnamed: 8': 'YOUTUBE_Link',
                                          'Unnamed: 9': 'Opening_hour', 'Unnamed: 10': 'Closing_hour',
                                          'Unnamed: 11': 'Enrolment_Link', 'Unnamed: 12': 'About_Us_Info',
                                          'OTHER DETAILS': 'boarding_available', 'Unnamed: 15': 'course_start',
                                          'Unnamed: 16': 'Type', 'Unnamed: 17': 'Ttion_fees_p_plan',
                                          "FACILITY'S & BORADING INFO": 'facilities', 'Unnamed: 20': 'sport_facilities',
                                          'Unnamed: 21': 'boarding_facilities'}, inplace=True)

    # drop all unnamed columns
    DataFrame_institution = DataFrame_institution[DataFrame_institution.columns.drop(list(DataFrame_institution.filter(regex='Unnamed')))]

    # drop rows which values are all NaN
    DataFrame_institution = DataFrame_institution.dropna(how='all')

    # drop first row, index[0], does not provide valuable infomation
    DataFrame_institution = DataFrame_institution.drop(DataFrame_institution.index[0])

    # adding new columns into Institution dataframe
    DataFrame_institution.insert(2, "Accredited", "Yes")
    DataFrame_institution.insert(8, "Latitude", "")
    DataFrame_institution.insert(9, "Longitude", "")
    DataFrame_institution.insert(11, "Schlr_finan_asst", "")
    DataFrame_institution.insert(23, "IMG_COUNT", "")

    # Course
    # columns are renamed to the provided final template, eg: from COUNTRY to Country
    # no need to rename certain unnamed columns to TERMLY domestic fees, international fees and boarding fees due to final excel template
    DataFrame.rename(columns={' ': 'Unnamed: 2', 'INSTITUTION NAME': 'University',
                              'CITY/TOWN': 'City', 'STATE/PROVINCE/ COUNTY': 'Province ',
                              'POSTCODE': 'post code', 'COUNTRY': 'Country',
                              'COURSE NAME': 'Courses', 'LANGUAGE OF INSTRUCTION': 'course_lang',
                              'COURSE LINK': 'Website', 'DOMESTIC FEES': 'Local_Fees', 'INTERNATIONAL FEES': 'Int_Fees',
                              'APPLICATION FEE': 'application_fe', 'ENROLMENT FEE': 'enrolment_fee',
                              'BOARDING FEE': 'boarding_fee', 'CURRENCY': 'Currency',
                              'ENTRANCE EXAM': 'enterance_exam', 'Remarks_': 'REMARKS',
                              'EDUCATION LEVEL': 'Course_Type'}, inplace=True)

    # drop all unnamed columns
    DataFrame = DataFrame[DataFrame.columns.drop(list(DataFrame.filter(regex='Unnamed')))]

    # drop columns which are not needed, curriculum, examination board, course_description (do not drop)
    # DataFrame = DataFrame.drop(columns=['CURRICULUM', 'EXAMINATION BOARD', 'COURSE DESCRIPTION'])

    # drop first row, does not contain any useful information
    DataFrame = DataFrame.drop(DataFrame.index[0])

    # drop rows which values are all NaN
    DataFrame = DataFrame.dropna(how='all')

    # inserting new columns
    DataFrame.insert(0, "availbilty", "All")
    DataFrame.insert(9, "Faculty", "")
    DataFrame.insert(14, "Currency_Time", "Year")
    DataFrame.insert(15, "Duration", "1")
    DataFrame.insert(16, "Duration_Time", "Year")

    # replacing course_type values with acronym, as required in final template
    DataFrame.loc[DataFrame['Course_Type'] == 'Early Years ', 'Course_Type'] = 'EY'
    DataFrame.loc[DataFrame['Course_Type'] == 'Primary (Elementary)', 'Course_Type'] = 'PRI'
    DataFrame.loc[DataFrame['Course_Type'] == 'Secondary (Middle School)', 'Course_Type'] = 'SECMS'
    DataFrame.loc[DataFrame['Course_Type'] == 'Secondary (High School)', 'Course_Type'] = 'SECHS'

    final_institution_data = final_institution_data.append(DataFrame_institution, ignore_index=True, sort=True)
    final_course_data = final_course_data.append(DataFrame, ignore_index=True, sort=True)

    print("Successfully added " + fileName +" to final list")


final_institution_data = final_institution_data.drop_duplicates()
final_institution_data = final_institution_data[final_institution_data.columns.drop(list(final_institution_data.filter(regex='Unnamed')))]

# rearranging columns to the final template
final_column_index_order = [16, 4, 1, 8, 7, 17, 13, 9, 10, 2, 12, 15, 18, 14, 21, 19, 22, 23, 20, 11, 3, 5, 0, 6]
final_institution_data = final_institution_data[[final_institution_data.columns[i] for i in final_column_index_order]]

final_course_data = final_course_data.drop_duplicates()
final_course_data = final_course_data[final_course_data.columns.drop(list(final_course_data.filter(regex='Unnamed')))]

# rearranging columns to the final template
final_column_index_order_2 = [19, 4, 16, 2, 14, 24, 3, 17, 5, 0, 1, 11, 10, 21, 13, 12, 6, 7, 8, 9, 18, 22, 20, 23, 15]
final_course_data = final_course_data[[final_course_data.columns[i] for i in final_column_index_order_2]]

# sort values
final_course_data.sort_values(by=['Country', 'University'])
final_institution_data.sort_values(by=['Country', 'Univeristy_Name'])

# export as xlsx file and close the program
final_course_data.to_excel(final_course_file_path, engine="xlsxwriter")
final_institution_data.to_excel(final_institution_file_path, engine="xlsxwriter")

stop = timeit.default_timer()
time_sec = stop - start
time_min = int(time_sec / 60)
time_hour = int(time_min / 60)

time_run = str(format(time_hour, "02.0f")) + ':' + str(
    format((time_min - time_hour * 60), "02.0f") + ':' + str(format(time_sec - (time_min * 60), "^-05.1f")))
print("This code has completed running in: " + time_run)




