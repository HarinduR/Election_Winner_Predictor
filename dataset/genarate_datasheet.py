import pandas as pd
import random

# Load the existing Excel file
file_path = 'dataset.xlsx'

# Load existing data (assuming it is empty now but has headers)
df = pd.read_excel(file_path)

# Number of voters to simulate
num_voters = 10000  # Change this value based on your needs

# List of districts (example list, add your own)
districts = [
    'Colombo', 'Gampaha', 'Kalutara', 'Kandy', 'Matale', 'Nuwara Eliya', 'Galle', 'Matara', 'Hambantota',
    'Jaffna', 'Kilinochchi', 'Mannar', 'Mullaitivu', 'Vavuniya', 'Trincomalee', 'Batticaloa', 'Ampara', 'Badulla',
    'Monaragala', 'Ratnapura', 'Kegalle', 'Puttalam', 'Kurunegala', 'Anuradhapura', 'Polonnaruwa'
]

# Define Nominator details with corresponding IDs
nominators = {1: 'Anura Kumara', 2: 'Sajith Premadasa', 3: 'Ranil Wickremasinghe', 4: 'Other'}

# Probability distribution for Nominators
nominator_probabilities = [0.51, 0.33, 0.15, 0.01]  # 51% for Anura, 43% Sajith, 5% Ranil, 1% Other

# Function to assign Age Group based on Age
def assign_age_group(age):
    if 18 <= age <= 25:
        return '18-25'
    elif 25 < age <= 35:
        return '25-35'
    elif 35 < age <= 55:
        return '35-55'
    elif 55 < age <= 75:
        return '55-75'
    else:
        return '75+'

# Generate the data
data = []
for i in range(1, num_voters + 1):
    age = random.randint(18, 100)
    district = random.choice(districts)
    nominator_id = random.choices([1, 2, 3, 4], nominator_probabilities)[0]
    nominator = nominators[nominator_id]
    age_group = assign_age_group(age)
    
    data.append([i, age, district, nominator_id, nominator, age_group])

# Create a DataFrame
new_data = pd.DataFrame(data, columns=['Number', 'Age', 'District', 'Nominator ID', 'Nominator', 'Age Group'])

# Write back to the Excel file without deleting the existing data
with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
    new_data.to_excel(writer, index=False)

print(f"Data successfully written to {file_path}")
