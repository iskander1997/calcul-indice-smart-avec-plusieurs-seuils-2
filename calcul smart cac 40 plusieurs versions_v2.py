###
# calcul de l'indice smart CAC40 avec plusieurs seuils de scores. 
# input : excels de cours de toutes les sociétés pour toutes les versions et excel de score smart smi
# output : dataframe avec les variations de l'indice smart cac40 pour chaque seuil
# v3 compare plusieurs scores "
##
import pandas as pd
import numpy as np
import logging
from Interface_selection_excel_et_scores_v2 import ExcelFileSelector 


def load_data(prices_path, scores_path):
    """
    Load stock prices and scores data from Excel files.
    
    Parameters:
    - prices_path: Path to the stock prices Excel file
    - scores_path: Path to the scores Excel file
    
    Returns:
    - Tuple of (prices_data, scores_data)
    """
    try:
        prices_data = pd.read_excel(prices_path, sheet_name="stock_prices")
        scores_data = pd.read_excel(scores_path, sheet_name="Versions")
        
        # Ensure Date columns are in datetime
        prices_data['Date'] = pd.to_datetime(prices_data['Date'])
        scores_data['Date'] = pd.to_datetime(scores_data['Date'], dayfirst=True)
        
        return prices_data, scores_data
    except Exception as e:
        logging.error(f"Error loading data : {e}")
        raise

def calculate_ponderation(scores_df, seuil):
    """
    Calculate weighting based on scores above a given threshold.
    
    Parameters:
    - scores_df: DataFrame with scores
    - seuil: Threshold for score-based weighting
    
    Returns:
    - DataFrame with added Difference and Ponderation columns
    """
    # Ensure Date is in datetime format
    scores_df = scores_df.copy()
    scores_df["Date"] = pd.to_datetime(scores_df["Date"], dayfirst=True)
    
    # Calculate the difference (only for scores >= seuil, else 0)
    scores_df["Difference"] = scores_df["SCORE"].apply(lambda x: x - seuil if x >= seuil else 0)
    
    # Calculate sum of differences for each Date
    sum_differences = scores_df.groupby("Date")["Difference"].transform("sum")
    
    # Compute ponderation
    scores_df["Ponderation"] = scores_df["Difference"] / sum_differences
    
    return scores_df

def clean_price_data(prices_data):
    """
    Cleans a price dataframe by filling missing values with the previous day's price.
    
    Parameters:
    - prices_data: DataFrame containing stock prices
    
    Returns:
    - Cleaned DataFrame with filled missing values
    """
    # Make a copy to avoid modifying the original
    prices_data_clean = prices_data.copy()
    
    # Ensure the DataFrame is sorted by date
    prices_data_clean = prices_data_clean.sort_values('Date')
    
    # Get all columns except 'Date'
    price_columns = [col for col in prices_data_clean.columns if (col != 'Date') & (col != 'CAC 40')]
    
    # For each price column, forward fill missing values
    for column in price_columns:
        # Forward fill (using previous value)
        prices_data_clean[column] = prices_data_clean[column].fillna(method='ffill')
    
    return prices_data_clean

def calculate_complete_smart_cac(prices_df, scores_df, seuil=125, verbose=True):
    """
    Calculate the SMART CAC40 index with comprehensive tracking and analysis.
    
    Parameters:
    - prices_df: DataFrame with stock prices
    - scores_df: DataFrame with company scores
    - seuil: Threshold for score-based weighting (default 125)
    - verbose: If True, prints detailed logging information
    
    Returns:
    - Dictionary with detailed calculation results
    """
    # Setup logging
    logging.basicConfig(
        level=logging.INFO if verbose else logging.WARNING,
        format='%(asctime)s - %(levelname)s: %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )
    logger = logging.getLogger(__name__)
    
    logger.info(f"Starting SMART CAC40 Calculation with Threshold: {seuil}")
    
    # Prepare score weighting
    try:
        scores_with_pond = calculate_ponderation(scores_df, seuil)
        version_dates = np.sort(scores_with_pond['Date'].unique())
        
        logger.info(f"Number of Versions Detected: {len(version_dates)}")
    except Exception as e:
        logger.error(f"Error preparing version dates: {e}")
        raise
    
    # Track companies in each version
    version_companies = {}
    for version_date in version_dates:
        current_scores = scores_with_pond[scores_with_pond['Date'] == version_date]
        high_score_companies = current_scores[current_scores['SCORE'] >= seuil]
        
        companies = high_score_companies['SYMBOLE'].tolist()
        companies_details = high_score_companies[['SYMBOLE', 'SCORE', 'Ponderation']].to_dict('records')
        
        # Print companies for each version
        print(f"\nCompanies in version {version_date} (Threshold {seuil}):")
        for company in companies_details:
            print(f"Symbol: {company['SYMBOLE']}, Score: {company['SCORE']}, Weight: {company['Ponderation']:.4f}")
        
        version_companies[str(version_date)] = {
            'symbols': companies,
            'total_count': len(companies),
            'companies_details': companies_details
        }
    
    # Prepare result DataFrame
    all_dates = prices_df['Date'].sort_values().unique()
    result_df = pd.DataFrame({'Date': all_dates})
    result_df['CAC 40'] = result_df['Date'].map(dict(zip(prices_df['Date'], prices_df['CAC 40'])))
    result_df['SMART CAC40'] = 0.0
    result_df['Version'] = None
    result_df['Total_Variation'] = 0.0
    
    # Assign version dates
    for i, date in enumerate(version_dates):
        if i < len(version_dates) - 1:
            mask = (result_df['Date'] >= date) & (result_df['Date'] < version_dates[i+1])
        else:
            mask = result_df['Date'] >= date
        result_df.loc[mask, 'Version'] = date
    
    result_df = result_df.sort_values('Date').reset_index(drop=True)
    
    # Core calculation
    last_smart_cac = None
    for version_date in version_dates:
        logger.info(f"Processing Version: {version_date}")
        
        # Filter rows for current version
        version_mask = result_df['Version'] == version_date
        version_rows = result_df[version_mask]
        
        if version_rows.empty:
            logger.warning(f"No data for version {version_date}")
            continue
        
        # Get high-scoring symbols and their weights
        current_scores = scores_with_pond[scores_with_pond['Date'] == version_date]
        high_score_symbols = current_scores[current_scores['SCORE'] >= seuil]['SYMBOLE'].tolist()
        ponderations = current_scores[current_scores['SYMBOLE'].isin(high_score_symbols)].set_index('SYMBOLE')['Ponderation']
        
        # Determine base value and reference date
        if last_smart_cac is None:
            base_value = result_df.loc[version_mask, 'CAC 40'].iloc[0]
            reference_date = version_rows['Date'].iloc[0]
        else:
            base_value = last_smart_cac
            prev_version_dates = result_df[result_df['Version'] < version_date]['Date']
            reference_date = prev_version_dates.max() if not prev_version_dates.empty else None
        
        # Get reference prices
        reference_prices = {}
        if reference_date is not None:
            ref_day_mask = prices_df['Date'] == reference_date
            if ref_day_mask.any():
                ref_day_row = prices_df[ref_day_mask].iloc[0]
                reference_prices = {symbol: ref_day_row[symbol] 
                                    for symbol in high_score_symbols 
                                    if symbol in prices_df.columns}
        
        # Calculate for each day in the version
        for idx, row in version_rows.iterrows():
            current_date = row['Date']
            current_day_mask = prices_df['Date'] == current_date
            
            if not current_day_mask.any():
                continue
            
            current_day_row = prices_df[current_day_mask].iloc[0]
            
            # Calculate weighted total variation
            total_var = 0
            for symbol in high_score_symbols:
                if symbol not in prices_df.columns or symbol not in reference_prices:
                    continue
                
                current_price = current_day_row[symbol]
                reference_price = reference_prices[symbol]
                
                # Skip invalid price calculations
                if pd.isna(reference_price) or pd.isna(current_price) or reference_price == 0:
                    continue
                
                # Calculate variation and weighted contribution
                variation = (current_price / reference_price - 1) * 100
                weighted_variation = variation * ponderations.get(symbol, 0)
                total_var += weighted_variation
            
            # Update results
            result_df.loc[idx, 'Total_Variation'] = total_var
            result_df.loc[idx, 'SMART CAC40'] = base_value * (1 + total_var / 100)
        
        # Update last SMART CAC40 value for next iteration
        if not version_rows.empty:
            last_smart_cac = result_df.loc[version_rows.index[-1], 'SMART CAC40']
    
    # Final DataFrame and calculations
    final_df = result_df[['Date', 'CAC 40', 'SMART CAC40', 'Total_Variation']]
    
    smart_cac40_values = final_df['SMART CAC40']
    first_smart_cac = final_df['SMART CAC40'].iloc[0]
    last_smart_cac = final_df['SMART CAC40'].iloc[-1]
    total_period_variation = (last_smart_cac / first_smart_cac - 1) * 100
    
    # Logging final results
    logger.info(f"smart_cac40_values: {smart_cac40_values}")
    logger.info(f"First SMART CAC40 Value: {first_smart_cac:.4f}")
    logger.info(f"Last SMART CAC40 Value: {last_smart_cac:.4f}")
    logger.info(f"Total Period Variation: {total_period_variation:.4f}%")
    
    return {
        'dataframe': final_df,
        'seuil': seuil,
        'smart_cac40_values': smart_cac40_values,
        'total_period_variation': total_period_variation,
        'version_companies': version_companies
    }


IN_COLAB = False
try:
    from google.colab import files
    IN_COLAB = True
except ModuleNotFoundError:
    pass

def main():
    """
    Main function to run the SMART CAC40 analysis with multiple threshold tests.
    """
    # Set up logging
    logging.basicConfig(level=logging.INFO, 
                        format='%(asctime)s - %(levelname)s: %(message)s',
                        datefmt='%Y-%m-%d %H:%M:%S')
    
    if IN_COLAB:
        print("Running in Google Colab environment")
        print("For Colab usage, please use this notebook approach instead:")
        print("1. Import the script without running it: `import calcul_smart_cac_40_plusieurs_versions_v2`")
        print("2. Call the selector interactively: `selector = calcul_smart_cac_40_plusieurs_versions_v2.ExcelFileSelector()`")
        print("3. Display the UI: `data_paths = selector.run()`")
        print("4. After validation, run: `calcul_smart_cac_40_plusieurs_versions_v2.run_analysis(data_paths)`")
        return
    
    selector = ExcelFileSelector()
    data_paths = selector.run()

    if data_paths:
        run_analysis(data_paths)
    else:
        logging.error("No data paths selected.")
        return

def run_analysis(data_paths):
    """Function to process data once paths are selected"""
    prices_path = data_paths['prices_path']
    scores_path = data_paths['scores_path']
    thresholds = data_paths['thresholds']
    
    try:
        # Load data
        prices_data, scores_data = load_data(prices_path, scores_path)

        # Clean price data
        prices_data_clean = clean_price_data(prices_data)

        # Test multiple thresholds
        results = {}

        for seuil in thresholds:
            print(f"\n=== Analysis with Threshold {seuil} ===")
            result = calculate_complete_smart_cac(prices_data_clean, scores_data, seuil=seuil)
            results[seuil] = result

        # Compare results
        print("\n--- Threshold Comparison ---")
        for seuil, result in results.items():
            print(f"Threshold {seuil}: Total Period Variation = {result['total_period_variation']:.4f}%")
        
        return results

    except Exception as e:
        logging.error(f"An error occurred: {e}")
        return None

if __name__ == "__main__":
    main()

