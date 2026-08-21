from datetime import datetime

from openpyxl.descriptors import DateTime


def format_date() -> str:
    now = datetime.now()
    formatted_now = now.strftime("%B %d, %Y").upper()

    # Print results
    print(f'Current Date: "{formatted_now}"')
    return formatted_now

# --- VERIFICATION EXAMPLES ---
if __name__ == "__main__":
    # Test 1: No date provided (Uses current date)
    print(f'Default (Current): "{format_date()}"')


