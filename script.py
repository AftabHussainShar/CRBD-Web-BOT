import pandas as pd
import pyautogui
import time
import pyperclip
import re

excel_file = 'ListofVATRegPersons.xlsx'
df = pd.read_excel(excel_file)
column_data = df.iloc[:, 2]
column_data.pop(0)
time.sleep(2)

# Switch to the target application window
pyautogui.hotkey('alt', 'tab')
companies = []
turnovers = []
financial_years = []
i=0
for index, column in enumerate(column_data):
    try:
        i = i+1
        pyautogui.click(x=183, y=316)
        pyautogui.write(column, interval=0.1)
        try:
            image_location = pyautogui.locateOnScreen('search.png', confidence=0.8)
            if image_location is not None:
                x, y = pyautogui.center(image_location)
                pyautogui.click(x, y)
            else:
                print("Image 'search.png' not found on the screen.")
        except Exception as e:
            print(f"An error occurred while searching for 'search.png': {str(e)}")
        
        time.sleep(5)
        try:
            image_location = pyautogui.locateOnScreen('eye.png', confidence=0.8)
            if image_location is not None:
                x, y = pyautogui.center(image_location)
                pyautogui.click(x, y)
            else:
                print("Image 'eye.png' not found on the screen.")
        except Exception as e:
            print(f"An error occurred while searching for 'eye.png': {str(e)}")
        
        time.sleep(2)
        pyautogui.hotkey('ctrl', 'a')
        pyautogui.hotkey('ctrl', 'c')
        copied_text = pyperclip.paste()
        start_index = copied_text.find("PROFIT AND LOSS STATEMENT")
        end_index = copied_text.find("Less cost of Sales")
        
        if start_index != -1 and end_index != -1:
            desired_section = copied_text[start_index:end_index]
            financial_year_pattern = r"Financial Year Ended:\s*([\d/]+)"
            turnover_pattern = r"Turnover\s*([\d,]+)"
            financial_year_match = re.search(financial_year_pattern, desired_section)
            turnover_match = re.search(turnover_pattern, desired_section)
            
            if financial_year_match:
                financial_year = financial_year_match.group(1).strip() 
            else:
                financial_year = None
            
            if turnover_match:
                turnover = turnover_match.group(1).replace(",", "")  
            else:
                turnover = None 
                
            companies.append(column)  
            turnovers.append(turnover)
            financial_years.append(financial_year)
            
        else:
            companies.append(column)
            turnovers.append(None)
            financial_years.append(None)
            
        # Clear the search bar for the next iteration
        time.sleep(1)
        pyautogui.click(x=1288, y=153)
        time.sleep(1)
        pyautogui.click(x=409, y=296)
        time.sleep(2)
        pyautogui.hotkey('ctrl', 'a')
        pyautogui.hotkey('ctrl', 'x')
        
    except Exception as e:
        print(f"An error occurred: {str(e)}")
        companies.append(column)
        turnovers.append(None)
        financial_years.append(None)
    if i == 20:
        break
    
# Create DataFrame
data = {"Company": companies, "Turnover": turnovers, "Financial Year": financial_years}
df = pd.DataFrame(data)

# Save DataFrame to CSV
df.to_csv("output.csv", index=False)
print("CSV file saved successfully.")
