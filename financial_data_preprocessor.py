import json
import logging
from typing import Dict, List, Any, Optional
from pathlib import Path

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

class FinancialDataProcessor:
    def __init__(self, config_path: str = 'financial_data_config.json'):
        """
        Initialize the financial data processor with configurations.
        
        Args:
            config_path: Path to the configuration JSON file
        """
        self.config_path = config_path
        self.config = self._load_config()
        
    def _load_config(self) -> Dict:
        """Load and validate the configuration file."""
        try:
            if Path(self.config_path).exists():
                with open(self.config_path, 'r') as f:
                    return json.load(f)
            else:
                logger.warning(f"No configuration file found at {self.config_path}")
                return {}
        except json.JSONDecodeError as e:
            logger.error(f"Error parsing configuration file: {e}")
            return {}
            
    def _apply_field_override(self, financial: Dict[str, Any], field: str, value: Any) -> None:
        """
        Apply an override to a specific field, handling special cases like formulas.
        
        Args:
            financial: The financial statement dictionary to modify
            field: The field to override
            value: The override value or formula
        """
        if isinstance(value, str) and value.startswith('='):
            # Handle formula cases
            formula = value[1:]  # Remove the '=' prefix
            if formula in financial:
                financial[field] = financial[formula]
            else:
                logger.warning(f"Formula reference '{formula}' not found in financial data")
        else:
            # Direct value override
            financial[field] = value

    def process_financial_data(self, 
                             ticker: str, 
                             data: Dict[str, List[Dict[str, Any]]], 
                             statement_type: str) -> Dict[str, List[Dict[str, Any]]]:
        """
        Process financial data applying any configured overrides.
        
        Args:
            ticker: The stock ticker symbol
            data: The financial data dictionary containing statements
            statement_type: Type of statement ('income_statement', 'balance_sheet', or 'cash_flow')
            
        Returns:
            The processed financial data
        """
        if not data or 'financials' not in data:
            logger.warning(f"No financial data found for {ticker}")
            return data
            
        ticker = ticker.upper()
        ticker_config = self.config.get(ticker, {})
        statement_config = ticker_config.get(statement_type, {})
        
        if not statement_config:
            return data
            
        logger.info(f"Applying {ticker} specific rules for {statement_type}")
        
        # Apply overrides to each financial statement
        for financial in data['financials']:
            for field, override_value in statement_config.items():
                if field in financial:
                    self._apply_field_override(financial, field, override_value)
                    logger.debug(f"Applied override for {ticker} - {field}: {override_value}")
                else:
                    logger.warning(f"Field {field} not found in financial data for {ticker}")
                    
        return data

def process_financial_statements(ticker: str, ic_data: Dict = None, bs_data: Dict = None, cf_data: Dict = None) -> tuple:
    """
    Process multiple financial statements with configuration-based overrides.
    
    Args:
        ticker: The stock ticker symbol
        ic_data: Income statement data (optional)
        bs_data: Balance sheet data (optional)
        cf_data: Cash flow data (optional)
        
    Returns:
        Tuple of processed (income_statement, balance_sheet, cash_flow) data
    """
    processor = FinancialDataProcessor()
    
    # Process each statement type if provided
    processed_ic = processor.process_financial_data(ticker, ic_data, 'income_statement') if ic_data else None
    processed_bs = processor.process_financial_data(ticker, bs_data, 'balance_sheet') if bs_data else None
    processed_cf = processor.process_financial_data(ticker, cf_data, 'cash_flow') if cf_data else None
    
    return processed_ic, processed_bs, processed_cf

# Example usage:
if __name__ == "__main__":
    # Example data
    sample_ic_data = {
        "financials": [
            {
                "date": "2023-12-31",
                "symbol": "MA",
                "revenue": 25098000000,
                "costOfRevenue": 5000000000,
                "grossProfit": 20098000000
            }
        ]
    }
    
    # Process the financial statements
    processed_ic, _, _ = process_financial_statements("MA", ic_data=sample_ic_data)
    
    # Print results
    if processed_ic:
        print("Processed Income Statement:", json.dumps(processed_ic, indent=2))