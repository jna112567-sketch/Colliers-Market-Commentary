import pdfplumber
import re
import statistics

def extract_text_from_pdf(file_stream):
    """Extracts all text from a PDF file stream using pdfplumber."""
    text = ""
    try:
        with pdfplumber.open(file_stream) as pdf:
            for page in pdf.pages:
                page_text = page.extract_text()
                if page_text:
                    text += page_text + "\n"
    except Exception as e:
        print(f"Error reading PDF: {e}")
    return text

def parse_market_figures(text):
    """
    Uses Regex to find Vacancy Rate and Rent figures in English or Korean.
    Returns a dictionary with extracted floats (or None if not found).
    """
    results = {
        'vacancy_rate': None,
        'rent': None
    }
    
    # Regex for Vacancy (e.g., "Vacancy Rate: 2.5%", "공실률 3.1%")
    # Looks for 'vacancy' or '공실' followed by optional chars, then a number, then '%'
    vacancy_pattern = r'(?i)(?:vacancy|공실률|공실율)[^\d]{0,15}(\d+(?:\.\d+)?)\s*%'
    vacancy_matches = re.findall(vacancy_pattern, text)
    if vacancy_matches:
        # Take the first match or average them if multiple? Let's take the first reasonable one
        try:
            results['vacancy_rate'] = float(vacancy_matches[0])
        except ValueError:
            pass

    # Regex for Rent (e.g., "Face Rent: 120,000", "임대료: 135,000")
    # Looks for 'rent' or '임대료' followed by a number that might have commas
    rent_pattern = r'(?i)(?:face rent|rent|임대료)[^\d]{0,15}([\d]{2,3}(?:,[\d]{3})+)'
    rent_matches = re.findall(rent_pattern, text)
    if rent_matches:
        try:
            # Clean commas and take the first match
            clean_rent = rent_matches[0].replace(',', '')
            results['rent'] = float(clean_rent)
        except ValueError:
            pass
            
    return results

def process_multiple_reports(file_streams):
    """
    Processes multiple PDF files and averages the extracted metrics.
    """
    all_vacancies = []
    all_rents = []
    
    for stream in file_streams:
        text = extract_text_from_pdf(stream)
        figures = parse_market_figures(text)
        
        if figures['vacancy_rate'] is not None:
            all_vacancies.append(figures['vacancy_rate'])
        if figures['rent'] is not None:
            all_rents.append(figures['rent'])
            
    # Calculate averages
    final_vacancy = statistics.mean(all_vacancies) if all_vacancies else None
    final_rent = statistics.mean(all_rents) if all_rents else None
    
    return {
        'files_processed': len(file_streams),
        'estimated_vacancy_rate': round(final_vacancy, 2) if final_vacancy else None,
        'average_face_rent': round(final_rent, 0) if final_rent else None
    }
