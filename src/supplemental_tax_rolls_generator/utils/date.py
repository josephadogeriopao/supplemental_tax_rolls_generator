from datetime import datetime

# Option A: Get the current date and time
now = datetime.now()
formatted_now = now.strftime("%B %d, %Y").upper()

# Option B: Format a specific hardcoded date (e.g., April 1, 2025)
specific_date = datetime(2025, 4, 1)
formatted_specific = specific_date.strftime("%B %d, %Y").upper()

# Print results
print(f'Current Date: "{formatted_now}"')
print(f'Specific Date: "{formatted_specific}"')