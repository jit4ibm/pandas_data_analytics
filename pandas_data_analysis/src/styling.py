"""
Module for applying styling to DataFrames and exporting to Excel.
"""

import pandas as pd
from pathlib import Path
import logging
from typing import Optional

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('data_styling.log'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

class DataStyler:
    """
    A class to apply styling to DataFrames and export to Excel.
    """
    
    def __init__(self, data: Optional[dict] = None, csv_path: Optional[Path] = None):
        """
        Initialize the DataStyler with either data dictionary or CSV path.
        
        Args:
            data: Dictionary containing data (default None)
            csv_path: Path to CSV file (default None)
            
        Raises:
            ValueError: If neither data nor csv_path is provided
        """
        if data is None and csv_path is None:
            raise ValueError("Either data or csv_path must be provided")
            
        self.data = data
        self.csv_path = csv_path
        self.df = None
        
    def load_data(self) -> pd.DataFrame:
        """
        Load data from either the provided dictionary or CSV file.
        
        Returns:
            Loaded DataFrame
            
        Raises:
            ValueError: If data loading fails
        """
        try:
            if self.data is not None:
                logger.info("Creating DataFrame from provided data")
                self.df = pd.DataFrame(self.data)
            else:
                logger.info(f"Loading data from CSV: {self.csv_path}")
                self.df = pd.read_csv(self.csv_path)
                
            logger.info(f"Data loaded successfully. Shape: {self.df.shape}")
            return self.df
            
        except Exception as e:
            logger.error(f"Error loading data: {e}")
            raise ValueError(f"Data loading failed: {e}")
    
    def apply_styling(self) -> pd.io.formats.style.Styler:
        """
        Apply styling to the DataFrame.
        
        Returns:
            Styler object with applied styles
            
        Raises:
            ValueError: If data hasn't been loaded
        """
        if self.df is None:
            raise ValueError("Data not loaded. Call load_data() first.")
            
        try:
            logger.info("Applying styling to DataFrame")
            
            # Create a base styler
            styler = self.df.style
            
            # Apply salary highlighting if column exists
            if 'Salary' in self.df.columns:
                styler = styler.apply(self._highlight_salary, subset=['Salary'])
                logger.info("Applied salary highlighting")
            
            # Apply negative value highlighting to numeric columns
            numeric_cols = self.df.select_dtypes(include='number').columns
            for col in numeric_cols:
                styler = styler.apply(self._highlight_negative, subset=[col])
            logger.info("Applied negative value highlighting")
            
            return styler
            
        except Exception as e:
            logger.error(f"Error applying styling: {e}")
            raise ValueError(f"Styling failed: {e}")
    
    @staticmethod
    def _highlight_salary(s: pd.Series) -> list:
        """
        Highlight cells in Salary column with values above 50000.
        
        Args:
            s: Pandas Series (Salary column)
            
        Returns:
            List of style strings
        """
        return ['background-color: yellow' if v > 50000 else '' for v in s]
    
    @staticmethod
    def _highlight_negative(s: pd.Series) -> list:
        """
        Highlight cells with negative values.
        
        Args:
            s: Pandas Series (numeric column)
            
        Returns:
            List of style strings
        """
        return ['background-color: red' if v < 0 else '' for v in s]

def main():
    """Main execution function."""
    try:
        # Define paths
        output_dir = Path("data/output")
        output_dir.mkdir(parents=True, exist_ok=True)
        output_file = output_dir / "styled_result.xlsx"
        
        # Sample data
        data = {
            'Name': ['John', 'Jane', 'Bob', 'Mary'],
            'Age': [25, 30, 20, 35],
            'Salary': [50000, 60000, 45000, 70000]
        }
        
        # Initialize and process data
        styler = DataStyler(data=data)
        df = styler.load_data()
        
        # Apply styling
        styled_df = styler.apply_styling()
        
        # Save to Excel
        logger.info(f"Saving styled data to {output_file}")
        styled_df.to_excel(output_file, index=False)
        logger.info("Styling and export completed successfully")
        
    except Exception as e:
        logger.error(f"Processing failed: {e}")

if __name__ == "__main__":
    main()