import pandas as pd

# Bestandspad
file_name = "AwardGrantsOverview.xlsx"

# Inlezen van de Excel-bestand
data = pd.read_excel(file_name)

# Controleren of de benodigde kolommen bestaan
required_columns = ['Color', 'Type', 'Number']
if not all(col in data.columns for col in required_columns):
    raise ValueError(f"De Excel file moet de volgende kolommen bevatten: {', '.join(required_columns)}")

# Sorteren op Color, Type en Number
data_sorted = data.sort_values(by=['Color', 'Type', 'Number'], ascending=[True, True, True])

# Groeperen op Color en Type en het hoogste nummer vinden
grouped = data_sorted.groupby(['Color', 'Type'])
highest_numbers = grouped['Number'].max()

# Printen van de resultaten
print("Hoogste nummers voor elke Color-Type combinatie:\n")
for (color, type_), number in highest_numbers.items():
    print(f"\t{color}, {type_}, Hoogste Number: {number}")

# Optioneel: opslaan van de gesorteerde data in een nieuw Excel-bestand
data_sorted.to_excel("AwardGrantsOverview.xlsx", index=False)
print("\n\n> Gesorteerde data is opgeslagen in 'AwardGrantsOverview.xlsx'.")
