from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from openpyxl import Workbook
import requests
from bs4 import BeautifulSoup
import os
import time
import pandas as pd
import tkinter as tk
from tkinter import messagebox
from tkinter import ttk
import openpyxl


def generateSchedule():
    # Get value of checkbox for headless mode
    bool_headless = checkbox_headless_var.get()

    # Create the options for webdriver
    options = webdriver.ChromeOptions()

    # Set the headless mode
    if bool_headless==1:
        options.add_argument("--headless=new")

    # Create the service and driver
    service = Service(executable_path="chromedriver.exe")
    driver = webdriver.Chrome(service=service, options=options)

    # Get the values of the username and password
    username = text_username.get()
    password = text_password.get()

    # Check if the username and password are empty
    if username == "" or password == "":
        messagebox.showerror("Login Error", "Please enter both username and password")
    
    # Open the browser and login to the portal
    driver.maximize_window()
    driver.get("https://cx.usiu.ac.ke/ics")
    WebDriverWait(driver,60).until(
        EC.presence_of_element_located((By.ID, "userName")))
    element_user_name = driver.find_element(By.ID, "userName")
    element_user_name.send_keys(username)
    element_pass_word = driver.find_element(By.ID, "password")
    element_pass_word.send_keys(password + Keys.ENTER)

    # Validating the login
    try:
        element = WebDriverWait(driver, 2).until(
        EC.presence_of_element_located((By.PARTIAL_LINK_TEXT, "Student"))
    )
    except Exception as e:
        driver.close()
        messagebox.showerror("Login Error", "Incorrect username or password")

    # Navigating to the Student page
    student = driver.find_element(By.PARTIAL_LINK_TEXT, "Student")
    student.click()

    WebDriverWait(driver,60).until(
        EC.presence_of_element_located((By.PARTIAL_LINK_TEXT, "Course Schedule and Registration"))
    )

    # Navigating to the Course Schedule and Registration page
    course_schedule = driver.find_element(By.PARTIAL_LINK_TEXT, "Course Schedule and Registration")
    course_schedule.click()

    # Waiting for the Course Schedule and Registration page to load
    WebDriverWait(driver,60).until(
    EC.presence_of_element_located((By.ID, "pg1_V_lblAdvancedSearch"))
    )

    # Navigating to the Course Search page
    advanced_search = driver.find_element(By.ID, "pg1_V_lblAdvancedSearch")
    advanced_search.click()

    # Waiting for the Course Search page to load
    WebDriverWait(driver,60).until(
    EC.presence_of_element_located((By.ID, "pg0_V_btnSearch"))
    )

    # Changing to the required semester
    dropdown_semester = driver.find_element(By.ID, "pg0_V_ddlTerm")
    chosen_semester = combo_semester.get()
    select = Select(dropdown_semester)
    select.select_by_visible_text(chosen_semester)

    # Waiting for the page to load
    WebDriverWait(driver,60).until(
        EC.presence_of_element_located((By.ID, "pg0_V_btnSearch"))
    )

    # Changing to the required lecturer
    dropdown_lecturer = driver.find_element(By.ID, "pg0_V_ddlFaculty")
    chosen_lecturer = combo_lecturer.get()
    select = Select(dropdown_lecturer)
    select.select_by_visible_text(chosen_lecturer)

    # Extracting the correct days to search
    checked_days = []
    for i in range(len(checkbox_vars)):
        if checkbox_vars[i].get() == 1:
            if days[i] == "MW":
                checked_days.append("Mon")
                checked_days.append("Wed")
            elif days[i] == "TR":
                checked_days.append("Tue")
                checked_days.append("Thu")
            elif days[i] == "F":
                checked_days.append("Fri")
            elif days[i] == "S":
                checked_days.append("Sat")
            elif days[i] == "M":
                checked_days.append("Mon")
            elif days[i] == "T":
                checked_days.append("Tue")
            elif days[i] == "W":
                checked_days.append("Wed")
            else:
                checked_days.append("Thu")
    
    # Setting the correct days to search
    for day in checked_days:
        check_day = driver.find_element(By.ID, f"pg0_V_chk{day}")
        check_day.click()
    

    # Searching for the course

    # Extracting the course codes from the text box
    course_codes = text_course_code.get("1.0", tk.END).split(",")
    
    # Array to store the retreived course data
    excel_data = []
    for course_code in course_codes:

        # Searching for the course
        try:
            course_code_to_add_to_dropdown = driver.find_element(By.ID, "pg0_V_txtCourseRestrictor")
            course_code_to_add_to_dropdown.send_keys(course_code)
            search = driver.find_element(By.ID, "pg0_V_btnSearch")
            search.click()
            table = WebDriverWait(driver,60).until(
                EC.visibility_of_element_located((By.ID, "tableCourses"))
            )

            # This part deals with course searches that span over multiple pages
            try:
                show_all = driver.find_element(By.ID, "pg0_V_lnkShowAllBottom")
                if show_all.text == "Show All":
                    show_all.click()
                    table = WebDriverWait(driver,60).until(
                    EC.visibility_of_element_located((By.ID, "tableCourses"))
                )
            except Exception as e:
                pass

            # Extracting the data from the table and storing it in the excel_data array
            rows = table.find_elements(By.TAG_NAME, "tr")
            for row in rows:
                cells = row.find_elements(By.TAG_NAME, "td")
                for cell in cells:
                    excel_data.append(cell.text)

            # Removing empty strings from the excel_data array caused by requisite, note and other fields not required
            for i in range(excel_data.count('')):
                excel_data.remove('')
            
            # Clicking the search again button to clear the search box and search for the next course if needed
            search_again = driver.find_element(By.ID, "pg0_V_glbSearchAgain")
            search_again.click()
            WebDriverWait(driver,60).until(
                EC.presence_of_element_located((By.ID, "pg0_V_btnSearch"))
            )
            course_code_to_add_to_dropdown = driver.find_element(By.ID, "pg0_V_txtCourseRestrictor")
            course_code_to_add_to_dropdown.clear()

        # Handling the case where no data is found for the course for the specified criteria
        except (NoSuchElementException, TimeoutException):
            # Adding the error message to the excel_data array to show on the sheet
            error_array = ["No","data","found","for",course_code,"for","the","selected","criteria"]
            for i in error_array:
                excel_data.append(i)

            # Navigate to the courses page to start a new search
            element = driver.find_element(By.XPATH, '//*[@id="youAreHere"]/li[5]/a')
            element.click()
            
            # Clear the search box
            course_code_to_add_to_dropdown = driver.find_element(By.ID, "pg0_V_txtCourseRestrictor")
            course_code_to_add_to_dropdown.clear()

    # Writing the content to an excel file
    row_count = len(excel_data) / 9
    column_count = 9
    new_list = [[None]*column_count for _ in range(int(row_count))]
    for i,value in enumerate(excel_data):
        row_index = i // column_count
        column_index = i % column_count
        new_list[row_index][column_index] = value
    
    # Writing the data to an excel file
    df = pd.DataFrame(new_list,columns=['Course Code','Course Title','Capacity','Status','Instructor','Credit Hours','Semester','Start Date','End Date'])

    # Adding blank rows between different courses
    i = 0
    while i < len(df)-1:
        if pd.notna(df['Course Code'][i]) and pd.notna(df['Course Code'][i+1]) and df['Course Code'][i][:7] != df['Course Code'][i+1][:7]:
            empty = pd.Series([pd.NA] * len(df.columns), index=df.columns)
            df = pd.concat([df.iloc[:i+1], pd.DataFrame([empty]), df.iloc[i+1:]]).reset_index(drop=True)
            i += 1  # Skip the next row because it's the empty row we just inserted
        i += 1
    df.to_excel('Course_Schedule.xlsx', index=False)

    # Adjusting the width of the columns to maximize the content
    wb = openpyxl.load_workbook('Course_Schedule.xlsx')
    sheet = wb['Sheet1']

    for column in sheet.columns:
        max_length = 0
        column = [cell for cell in column]
        for cell in column:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(cell.value)
            except:
                pass
        adjusted_width = (max_length + 2)
        sheet.column_dimensions[column[0].column_letter].width = adjusted_width

    # Saving the file and opening it
    wb.save('Course_Schedule.xlsx')    
    os.startfile('Course_Schedule.xlsx')

    time.sleep(5)

    # Closing the browser
    driver.close()

# User interface creation
# Create the main window
root = tk.Tk()
root.title("USIU-Africa Course Schedule and Registration")

# Set the window's width and height
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()

# Calculate the window's width and height
window_width = 550
window_height = 800

# Calculate the x and y positions to center the window
x = (screen_width // 2) - (window_width // 2)
y = (screen_height // 2) - (window_height // 2)

# Set the window size and position
root.geometry(f"{window_width}x{window_height}+{x}+{y-30}")
root.resizable(False, False)

# Create the widgets and place them in the window (using grid layout manager)
# Create the labels
label_login = tk.Label(root, text="Login to the CX Portal", font=("Arial", 16, "bold"))
label_login.grid(row=0, column=0, columnspan=3, padx=10,pady=10,sticky='ew')

label_username = tk.Label(root, text="Username", font=("Arial", 12, "bold"))
label_username.grid(row=1, column=0, pady=10, padx=10, sticky='e')

text_username = tk.Entry(root,font=("Arial", 12))
text_username.grid(row=1, column=1, columnspan=1,padx=10, pady=10, sticky='w')

label_password = tk.Label(root, text="Password", font=("Arial", 12, "bold"))
label_password.grid(row=2, column=0, pady=10, padx=10, sticky='e')

text_password = tk.Entry(root, font=("Arial", 12), show="*")
text_password.grid(row=2, column=1, columnspan=1, padx=10,pady=10, sticky='w')

separator1 = ttk.Separator(root,orient='horizontal')
separator1.grid(row=3, column=0, columnspan=3, padx=10,pady=10,sticky='nsew')

label_course_search = tk.Label(root, text="Advanced Course Search", font=("Arial", 16, "bold"))
label_course_search.grid(row=4, column=0, columnspan=3, pady=10,padx=10, sticky='ew')

label_course_code = tk.Label(root, text="Course Code", font=("Arial", 12, "bold"))
label_course_code.grid(row=5, column=0, pady=7, padx=10, sticky='e')

text_course_code = tk.Text(root, width=40,height=5,font=("Arial", 12))
text_course_code.grid(row=5, column=1, columnspan=2, pady=10, padx=10, sticky='w',rowspan=3)

label_semester = tk.Label(root, text="Semester", font=("Arial", 12, "bold"))
label_semester.grid(row=8, column=0, pady=10, padx=10, sticky='e')

semesters = ["FS 2022","SS 2022","US 2022","FS 2023","SS 2023","US 2023","SS 2024","US 2024"]
combo_semester = ttk.Combobox(root, width=40,values=semesters, font=("Arial", 12), state="readonly")
combo_semester.grid(row=8, column=1,columnspan=2, pady=10, padx=10, sticky='w')
combo_semester.current(len(semesters)-1)

list_lecturer = ['All', 'Adar, Korwa G', 'Adema, Valerie Palapala', 'Afundi, Patrick Omuhinda', 'Agade Mkutu, Kennedy', 'Akosa, E. Wambalaba', 'Akuma, Samson Mainye', 'Akundabweni, Loreen', 'Albert, Josiah', 'Ali, Fatuma Ahmed', 'Aloo, Linus Alwal', 'Alukwe, Chrispus Akhonya', 'Aluoch, Constance', 'Amuhaya, Edith Khavwajira', 'Andango, Elizabeth', 'Arasa, Josephine', 'Barasa, Constantine Mulondanome', 'Basweti, Kevin Ogachi', 'Behr, Agnes Wanjiru', 'Bichanga, Lawrence Areba', 'Bii, Cosmas K', 'Biko, Stephen', 'Bironga, Sophia Moraa', 'Buyu, Matthew', 'Bwire, Albert C.', "Chang'orok, Susan", 'Chege, Gerald W.', 'Chemonges, Cynthia Chepkorir,, M', 'Cherutich, Isaiah Kibet', 'Defersha, Amsalu Degu', "Diang'a, Racheal,, Dr.", 'Fernandez, Miguel', 'Gachanga, Esther Wanjiru', 'Gachukia, Jennifer Wangari', 'Gatumo, Francis Mambo', 'Getecha, Ciru W.', 'Gitahi, Jesse Elikanah_Machira', 'Githaiga, Paul W,, Mr', 'Githinji, Keziah Wangui,, Ms.', 'Githinji, Stanley Muturi,, Dr.', 'Githiri, John Gitonga', 'Githua, Edgar', 'Gromov, Mikhail D.', "Hussain, Syeda Re'em", "Iraki, Fredrick Kang'ethe", 'James, Sylvester Mutua_Kisila', 'Janapati, Yasodha Krishna', 'Jefwa, Judith Jai_Jaleha', 'Joseph, Muchina', 'Juma, Brenda', "K'aol, George O.", 'Kaburu, Mercy Kathambi', 'Kahiri, Caroline Njeri,, Dr.', 'Kakai, Pius Wanyonyi', 'Kaluyu, Veronicah', 'Kamau, Esther Njambi', 'Kamau, Joseph', 'Kamau, Sheila Wanjiku', 'Kangu, Maureen', 'Karanja, Moses Mwangi,, Mr', 'Karanja, Richard,, Mr', 'Karimi, James Mark,, Ngari', 'Kariuki, Veronica Waithira', 'Karugu, Mureithi', 'Karuri, Fridah', 'Katuse, Paul', 'Kavoo, Mark Kavoo', 'Kayeyia, Ernest Madara', 'Khadioli, Nancy', 'Khamala, Martin_Aaron Nawayo', 'Khayundi, Francis', 'Kiaye, Mary A.', 'Kibuku, Racheal Njeri', 'Kidaha, Alfred Usagi', 'Kihara, Allan S,, N', 'Kihara, Michael', 'Kihara, Tabitha Muthoni,, Ms', 'Kiilu, Damiana M.', 'Kiiyukia, Ciira', 'Kikete, Siambi H,, DR', 'Kimani, Gabriel Mr', 'Kimani, John Wainaina', 'Kimani, Larry M', 'Kimathi, Caroline Kagwiria', 'Kimotho, Stephen Gichuhi', 'Kioko, Angelina Nduku', 'Kioni, Benson Muthoga', 'Kipyegon, Shadrack', 'Kiriri, Peter Ndungu', "Kirui, Gideon Kipng'eno", 'Koshal, Jeremiah Ntaloi_Ole', 'Kuria,, John', 'Linge, Teresia K.K.', 'Lio, Sammy A.N.', 'Macharia, Carolyne Njoki,, Ms', 'Macharia, Hannah Muthoni', 'Macharia, Jimmy', 'Magambo, John Odhiambo', 'Mage, Grace Nduta', 'Magut, Zuhra C.', 'Mahmoud, Hussein Abdullahi', 'Maina, Ann Wambui', 'Maina, Muchara Dr', 'Maina, William Waweru', 'Mairura, Christopher Joseph', 'Maiyo, Joshua K.', 'Makori, Wycliffe Arika,, Dr', 'Mandela, Japheth Isaboke', 'Maore, Stephen K.', 'Maumo, Leonard Oluoch', 'Mbae, Justus', 'Mbatia, Betty Nyambura', 'Mbiriri, Michael K.', 'Mbogo, Marion Njeri', 'Mbotu, Michael Mulandi', "Mbugua, Levi Ng'ang'a", 'Mbugua, Paul Mungai', 'Mbugua, Peter Getyngo', 'Mbugua, Samuel Mungai,, Dr', 'Mbugua, Wanjiku', 'Mburu, Martin', 'Mbutu, Paul Mutinda', 'Milton, Obote J.', 'Minde, Nicodemus Michael', 'Mirichii, John Mwaniki', 'Misawo, Florence Awuor', 'Mohamed, Abdullahi Mohamed', 'Mohamed, Hussein Abdi', 'Mohamed, Mwanashehe S.', 'Mohan, Vaishnavi Ram', 'Muasya, Jane Nzisa', 'Muchemi, Joyce Karungari', 'Muchiri, Catherine Wanjiru', 'Muchiri, Jane Wairimu', 'Muchwanju, Chris', 'Muendo, Daniel', 'Mugo, Nancy W._Ms', 'Muhanji, Clare Imbosa,, Dr.', 'Muhonza, Prescott', 'Muindi, Benjamin', 'Mulindi, Patrick M.', 'Mulinge, Munyae M.', 'Mulinya, Sheila Joy', 'Mulwa, Harrison Munyao', 'Munene, Karega', 'Munene, Macharia', 'Mungai, Nelly Nyambura', 'Munyae, Margaret M.', 'Munyendo, Lincoln Linus_Were', 'Munyithya, James Mwendwa', 'Munyoki, Janvan Nzamba,, Dr', 'Munywoki, Vincent M', 'Muraguri, Charity Wairimu', 'Muriithi, Jane Gathigia', 'Muriithi, Petronilla Muthoni', 'Muriu, Marylyn Doreen', 'Mursi, Japheth', 'Musa, Grace Akinyi', 'Musau, Josephine Ndanu', 'Musebe, Edward Achieng', 'Musuva, Gladys Mwende', 'Mutanu, Leah', 'Mutisya Mutungi, Mary Mumbua', 'Mutwiri, Marion Kirumba', 'Mwakina, Rozina Munna', 'Mwalili, Tobias Mbithi', 'Mwangi, Anne Wairimu', 'Mwangi, Cyrus Wanjohi', 'Mwangi, Eunice', 'Mwangi, Isaac', 'Mwangi, Johnson Muthii', 'Mwangi, Jonathan Maina', 'Mwangi, Peterson Kimiru', 'Nakamura, Katsuji', 'Namada, Juliana Mulaa', 'Ndegwa, Joyce Watetu', 'Ndemo, Kwansah', 'Ndero, Peter K.', 'Nderu, Lawrence', 'Ndiege, Joshua Rumo', 'Ndirangu, Dalton', 'Nduati, Gidraph J.', 'Ndungu, Samuel Ndiritu', 'Ndungu, Tabitha K.', 'Nerubucha, David Wafula', 'Newa, Elsie Opiyo', 'Newa, Fred Omondi', "Ng'ang'a, Lucy N.", 'Nganga, Loise Wanjiku', 'Ngarachu, Fiona W.', 'Ngeru, Geoffrey Gacemi', 'Ngesa, Maureen O.', 'Ngui, Thomas Katua', 'Ngure, Beatrice W.', 'Ngware, Stephen Githaiga', 'Njenga, Kefah Muiruri', 'Njeri, Teresia', 'Njeru, Godwin Kinyua', 'Njihia, David Thiru', 'Njogu, Valentine Nyokabi', 'Njoroge, Dorothy Wanjiku,, Dr.', 'Njoroge, Geoffrey Githu', 'Njoroge, Gladys Gakenia', 'Njoroge, Joseph', 'Njoroge, Margaret', 'Njoroge, Martin Chege', 'Njoroge, Simon Githaiga', 'Njuguna, Joseph Kimani', 'Njui, Francis', "Njung'e, Richard Kagia,, Dr.", 'Noah, Naumih M.', 'Ntabo, Victor O.', 'Nyagwencha, Justus Nyamweya', 'Nyagwencha, Stella', 'Nyakundi, Jane', 'Nyamasyo, Eunice A.', 'Nyambati, Grace K', 'Nyambegera, Stephen Morangi', 'Nyamu, David G,, Dr.', 'Nyangweso, Silvester Muni', 'Nyanjom, Steven Ger,, Dr', 'Nyanoti, Joseph Nyamwange', 'Nyaribo, Cyprian Mose', 'Nyarigoti, Naom Moraa', 'Nyariki, Eric Mogaka', 'Nyayieka, Moureen Adhiambo', 'Nyete, Abraham Mutunga', 'Obila, Onyango James,, Mr.', "Ochieng', Robi Koki", 'Ochieng, Lucy Atieno', 'Ochola, Phares B', 'Odek, Anthony W.', 'Odera, Austin Owuor', 'Odhiambo, Terry J.', 'Odoyo, Fredrick', 'Ogada, Agnes Owuor', 'Ogenga, Daniel Okumu', 'Ogore, Fredrick Michae', 'Ogunde, Jane  Awuor', 'Ogutu, James O', 'Okanda, Paul M', 'Okaru, Alex Ogero', 'Okech, Timothy C.', 'Okello, Gabriel', 'Oketch, Omondi', 'Oluoch, Fred Ochieng', "Oluoch-Suleh, Everlyn Achieng'", 'Omboi, Bernard Messah', 'Ombui, Edward,, Dr', 'Omollo, Richard Otieno', 'Omolo, Calvin Andeve', 'Omondi, Daniel Onyango', 'Omulo, Elisha Opiyo,, Dr', 'Ondiek, Collins Oduor', 'Ondieki, Benard Odhiambo', 'Onyancha, Jared Misonge', 'Onyango, Moses', 'Ooko, George Opondo,, Mr.', 'Ooko, Maureen Achieng', 'Opondo, Mary Awuor', 'Oriedi, David Opondo', 'Oteri, Malack Omae', 'Otiende, Verrah Akinyi', 'Otieno, Bernard', 'Otieno, Pauline Adhiambo', 'Ouma, Caren Akomo', 'Ouma, Judy Aluoch', 'Ouma, Zackayo Omolo', 'Ouma,, Duncan Ochieng', 'Owili, Florence Akinyi', 'Owino, Joseph Owuor', 'Owuor, Benard O', 'Owuor, John David_Ouma', 'Oyaro, Kepha N._Makori', 'Rao, Jonnalagadda Venkateswara', 'Rintari, Ann Wanjiku', 'Rono, Ruthie C.', 'Rubio Gijon, Pablo', 'Scott, Bellows J', 'Sifuna, Austin Makokha', 'Sikalieh, Damaris', 'Sikolia, Geoffrey Serede,, Dr.', 'Sirma, Julius', 'Staff, Faculty', 'Sule, Odhiambo F.E.', 'Sungi, Simeon P.', 'Terefe, Ermias Mergia', 'Thuo, Martha Wanjiku', 'Vungo, Lilian Munanie_Nzia', 'Wafula, Maurine Maraka', 'Waga, Fred', 'Wainaina, Samuel', 'Wairagu, Peninah Muthoni,, Dr', 'Waithima, Charity Wangui', 'Wamai, Njoki E.', 'Wamalwa, Moses Kevin', 'Wambalaba, Francis', 'Wamuyu, Patrick Kanyi', 'Wangai, Njoroge Mambo', 'Wangechi, Chege', 'Watson, Carol J.', 'Were, Jamen H.', 'Were, Jane N.', 'Wokabi, Francis', 'Ylonen, Aleksi Erik']

label_lecturer = tk.Label(root, text="Lecturer", font=("Arial", 12, "bold"))
label_lecturer.grid(row=9, column=0, pady=10, padx=10, sticky='e')

combo_lecturer = ttk.Combobox(root, width=40,values=list_lecturer, font=("Arial", 12), state="readonly")
combo_lecturer.grid(row=9, column=1,columnspan=2, pady=10, padx=10, sticky='w')
combo_lecturer.current(0)


label_days = tk.Label(root, text="Days", font=("Arial", 12, "bold"))
label_days.grid(row=10, column=0, pady=10, padx=10, sticky='e')
checkboxes = []
checkbox_vars = []
days = ['MW', 'TR', 'F', 'S', 'M', 'T', 'W', 'R']
for i in range(8):
    checkbox_var = tk.IntVar(value=0)
    checkbox = tk.Checkbutton(root, width=5,text=f"{days[i]}", font=("Arial", 12), variable=checkbox_var)
    if i < 4:
        checkbox.grid(row=i+10, column=1, pady=10, padx=10, sticky='w')
    else:
        checkbox.grid(row=i+6, column=2, pady=10, padx=10, sticky='w')
    checkboxes.append(checkbox)
    checkbox_vars.append(checkbox_var)

separator2 = ttk.Separator(root,orient='horizontal')
separator2.grid(row=14, column=0, columnspan=3, padx=10,pady=10,sticky='nsew')

checkbox_headless_var = tk.IntVar(value=0)
checkbox_headless = tk.Checkbutton(root, text="Use Headless mode", font=("Arial", 12), variable=checkbox_headless_var)
checkbox_headless.grid(row=15, column=1, pady=10, padx=10, sticky='nsew')

button_generate = tk.Button(root,width=12 ,text="Generate Schedule", font=("Arial", 12, "bold"), bg="lightgreen", fg="black",command=generateSchedule)
button_generate.grid(row=16, column=1, pady=10, padx=10, sticky='nsew')

label_copyright = tk.Label(root, text="©️2024 - Developed by Shaurav Vora", font=("Arial", 10), fg="grey")
label_copyright.grid(row=17, column=0, columnspan=3, pady=10, padx=10, sticky='ew')

# Run the main loop
root.mainloop()