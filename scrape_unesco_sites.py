import requests
from bs4 import BeautifulSoup
import pandas as pd
import time
from typing import List, Dict

def scrape_unesco_sites() -> List[Dict]:
    """
    Scrapes all UNESCO World Heritage Sites from the official UNESCO website.
    Returns a list of dictionaries containing site information.
    """
    
    base_url = "https://whc.unesco.org/en/list/"
    all_sites = []
    page = 1
    
    print("Starting to scrape UNESCO World Heritage Sites...")
    
    while True:
        try:
            # Construct URL with pagination
            url = f"{base_url}?search=&order=&mode=table&action=list"
            if page > 1:
                url += f"&start={(page-1)*20}"
            
            print(f"Fetching page {page}...")
            
            headers = {
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
            }
            
            response = requests.get(url, headers=headers, timeout=10)
            response.raise_for_status()
            
            soup = BeautifulSoup(response.content, 'html.parser')
            
            # Find all site rows in the table
            rows = soup.find_all('tr')[1:]  # Skip header row
            
            if not rows:
                print("No more rows found. Scraping complete.")
                break
            
            for row in rows:
                cols = row.find_all('td')
                if len(cols) >= 4:
                    try:
                        # Extract site information
                        site_name = cols[0].text.strip()
                        country = cols[1].text.strip()
                        site_type = cols[2].text.strip()
                        year_inscribed = cols[3].text.strip()
                        
                        # Try to get criteria if available
                        criteria = cols[4].text.strip() if len(cols) > 4 else "N/A"
                        
                        site_data = {
                            'Site Name': site_name,
                            'Country': country,
                            'Type': site_type,
                            'Year Inscribed': year_inscribed,
                            'Criteria': criteria
                        }
                        
                        all_sites.append(site_data)
                    except Exception as e:
                        print(f"Error processing row: {e}")
                        continue
            
            page += 1
            time.sleep(1)  # Be respectful to the server
            
        except requests.exceptions.RequestException as e:
            print(f"Error fetching page {page}: {e}")
            break
        except Exception as e:
            print(f"Unexpected error on page {page}: {e}")
            break
    
    return all_sites


def scrape_unesco_sites_advanced() -> List[Dict]:
    """
    Alternative method using UNESCO's API or direct data access.
    This method attempts to get more detailed information.
    """
    
    all_sites = []
    
    try:
        # Try to use UNESCO's API endpoint if available
        api_url = "https://whc.unesco.org/en/list/json/"
        
        print("Attempting to fetch UNESCO sites from API...")
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
        }
        
        response = requests.get(api_url, headers=headers, timeout=10)
        response.raise_for_status()
        
        data = response.json()
        
        if 'results' in data:
            for site in data['results']:
                site_data = {
                    'Site Name': site.get('name', 'N/A'),
                    'Country': site.get('states', 'N/A'),
                    'Type': site.get('category', 'N/A'),
                    'Year Inscribed': site.get('date_inscribed', 'N/A'),
                    'Criteria': site.get('criteria', 'N/A'),
                    'Latitude': site.get('latitude', 'N/A'),
                    'Longitude': site.get('longitude', 'N/A'),
                    'Region': site.get('region', 'N/A')
                }
                all_sites.append(site_data)
        
        print(f"Successfully retrieved {len(all_sites)} sites from API")
        return all_sites
        
    except Exception as e:
        print(f"API method failed: {e}. Falling back to web scraping...")
        return scrape_unesco_sites()


def save_to_csv(sites: List[Dict], filename: str = 'unesco_world_heritage_sites.csv'):
    """
    Saves the sites data to a CSV file.
    """
    if not sites:
        print("No sites to save!")
        return
    
    df = pd.DataFrame(sites)
    df.to_csv(filename, index=False, encoding='utf-8')
    print(f"✓ Saved {len(sites)} sites to {filename}")


def save_to_excel(sites: List[Dict], filename: str = 'unesco_world_heritage_sites.xlsx'):
    """
    Saves the sites data to an Excel file.
    """
    if not sites:
        print("No sites to save!")
        return
    
    df = pd.DataFrame(sites)
    
    with pd.ExcelWriter(filename, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='All Sites', index=False)
        
        # Add formatting
        worksheet = writer.sheets['All Sites']
        for column in worksheet.columns:
            max_length = 0
            column_letter = column[0].column_letter
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = min(max_length + 2, 50)
            worksheet.column_dimensions[column_letter].width = adjusted_width
    
    print(f"✓ Saved {len(sites)} sites to {filename}")


def display_summary(sites: List[Dict]):
    """
    Displays a summary of the collected data.
    """
    if not sites:
        print("No data to display!")
        return
    
    df = pd.DataFrame(sites)
    
    print("\n" + "="*80)
    print(f"TOTAL UNESCO WORLD HERITAGE SITES: {len(sites)}")
    print("="*80)
    
    print("\nSites by Type:")
    print(df['Type'].value_counts())
    
    print("\nTop 15 Countries with Most Sites:")
    print(df['Country'].value_counts().head(15))
    
    print("\nFirst 10 Sites:")
    print(df.head(10).to_string())
    
    print("\n" + "="*80)


if __name__ == "__main__":
    # Scrape all UNESCO sites
    sites = scrape_unesco_sites_advanced()
    
    # Display summary
    display_summary(sites)
    
    # Save to files
    if sites:
        save_to_csv(sites)
        save_to_excel(sites)
        print(f"\n✓ Successfully created database with {len(sites)} UNESCO World Heritage Sites!")
    else:
        print("Failed to retrieve any sites. Please check your internet connection.")
