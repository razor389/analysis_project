import json
import argparse
import logging
from typing import Dict, List, Any

# Configure logging
logger = logging.getLogger(__name__)

def setup_logging(debug: bool = False):
    """
    Set up logging configuration.
    
    Args:
        debug: If True, set logging level to DEBUG, otherwise INFO
    """
    logging_level = logging.DEBUG if debug else logging.INFO
    
    logging.basicConfig(
        level=logging_level,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )

def transform_sec_data(input_data: Dict[str, Any], config: Dict[str, Any]) -> Dict[str, Dict[str, Dict[str, float]]]:
    """
    Transform SEC data according to the provided configuration.
    """
    result = {}
    years_data = input_data.get("years", {})
    
    for year, year_data in years_data.items():
        result[year] = {
            "revenue": {},
            "operating_income": {},
            "segmentation": {}
        }
        
        for category, category_config in config.items():
            target_tag = category_config["tag"]
            required_axes = set(category_config["axes"])
            name_mapping = category_config["name_mapping"]
            
            category_data = year_data.get(category, [])
            if not category_data and category == "segmentation":
                category_data = year_data.get("revenue", [])
            
            for entry in category_data:
                # Check tag match
                if entry["tag"] != target_tag:
                    continue
                
                # Check if ANY required axis is present (not all)
                entry_axes = set(entry["axis"].split("\n"))
                if not any(axis in entry_axes for axis in required_axes):
                    continue
                
                # Order-insensitive member matching
                entry_member_parts = set(entry["explicit_member"].split("\n"))
                for config_member, mapped_name in name_mapping.items():
                    config_member_parts = set(config_member.split("\n"))
                    if entry_member_parts == config_member_parts:
                        result[year][category][mapped_name] = entry["fact"]
                        break
    
    return result

def load_config(ticker: str) -> Dict[str, Any]:
    """
    Load configuration for the specified ticker from the config file.
    
    Args:
        ticker: Stock ticker symbol
    
    Returns:
        Configuration dictionary for the ticker
    """
    logger.info(f"Loading configuration for ticker {ticker}")
    try:
        with open('segmentation_transformation_config.json', 'r') as f:
            config = json.load(f)
    except FileNotFoundError:
        logger.error("Configuration file 'segmentation_transformation_config.json' not found")
        raise FileNotFoundError("Configuration file 'segmentation_transformation_config.json' not found")
    except json.JSONDecodeError:
        logger.error("Invalid JSON in configuration file 'segmentation_transformation_config.json'")
        raise ValueError("Invalid JSON in configuration file 'segmentation_transformation_config.json'")
    
    if ticker not in config:
        logger.error(f"Configuration not found for ticker: {ticker}")
        raise ValueError(f"Configuration not found for ticker: {ticker}")
    
    # Update the axes name to match the data
    if "segmentation" in config[ticker]:
        axes = config[ticker]["segmentation"]["axes"]
        for i, axis in enumerate(axes):
            if axis == "Product Or Services Axis":
                axes[i] = "Product Or Service Axis"
                logger.debug("Updated 'Product Or Services Axis' to 'Product Or Service Axis'")
    
    logger.debug(f"Configuration for {ticker}: {json.dumps(config[ticker], indent=2)}")
    return config[ticker]

def main():
    # Set up argument parser
    parser = argparse.ArgumentParser(description='Transform SEC data for a specific ticker and year')
    parser.add_argument('ticker', help='Stock ticker symbol')
    parser.add_argument('year', help='Year to process')
    parser.add_argument('--debug', action='store_true', help='Enable debug logging')
    args = parser.parse_args()

    # Setup logging
    setup_logging(args.debug)
    
    # Construct input filename
    input_filename = f"{args.ticker}_{args.year}_historical_breakdown.json"
    output_filename = f"{args.ticker}_{args.year}_segmentation.json"

    try:
        logger.info(f"Starting processing for {args.ticker} year {args.year}")
        
        # Load input data
        logger.info(f"Reading input file: {input_filename}")
        with open(input_filename, 'r') as f:
            input_data = json.load(f)
        
        # Load configuration for the ticker
        config = load_config(args.ticker)
        
        # Transform the data
        logger.info("Transforming data")
        transformed_data = transform_sec_data(input_data, config)
        
        # Save the transformed data
        logger.info(f"Saving results to: {output_filename}")
        with open(output_filename, 'w') as f:
            json.dump(transformed_data, f, indent=2)
            
        logger.info("Processing completed successfully")
        
    except FileNotFoundError as e:
        logger.error(f"File not found: {str(e)}")
        raise
    except json.JSONDecodeError as e:
        logger.error(f"Invalid JSON: {str(e)}")
        raise
    except Exception as e:
        logger.error(f"Unexpected error: {str(e)}")
        raise

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        logger.error("Program terminated with error", exc_info=True)
        exit(1)