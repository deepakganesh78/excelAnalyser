import pandas as pd
import numpy as np
from scipy import stats

class DataAnalyzer:
    def __init__(self, data):
        self.data = data
    
    def get_basic_info(self):
        """Get basic information about the dataset."""
        return {
            "Total Rows": len(self.data),
            "Total Columns": len(self.data.columns),
            "Memory Usage (MB)": round(self.data.memory_usage(deep=True).sum() / 1024 / 1024, 2),
            "Numerical Columns": len(self.data.select_dtypes(include=[np.number]).columns),
            "Categorical Columns": len(self.data.select_dtypes(include=['object']).columns),
            "DateTime Columns": len(self.data.select_dtypes(include=['datetime64']).columns)
        }
    
    def get_column_info(self):
        """Get detailed information about each column."""
        info_data = []
        for col in self.data.columns:
            info_data.append({
                "Column": col,
                "Data Type": str(self.data[col].dtype),
                "Non-Null Count": self.data[col].count(),
                "Null Count": self.data[col].isnull().sum(),
                "Unique Values": self.data[col].nunique()
            })
        return pd.DataFrame(info_data)
    
    def get_summary_statistics(self):
        """Get summary statistics for numerical columns."""
        numerical_data = self.data.select_dtypes(include=[np.number])
        if not numerical_data.empty:
            return numerical_data.describe()
        return pd.DataFrame()
    
    def get_data_quality_metrics(self):
        """Calculate various data quality metrics."""
        total_cells = self.data.shape[0] * self.data.shape[1]
        missing_cells = self.data.isnull().sum().sum()
        duplicate_rows = self.data.duplicated().sum()
        
        return {
            "Total Cells": total_cells,
            "Missing Cells": missing_cells,
            "Missing Percentage": round((missing_cells / total_cells) * 100, 2),
            "Duplicate Rows": duplicate_rows,
            "Duplicate Percentage": round((duplicate_rows / len(self.data)) * 100, 2),
            "Data Completeness": round(((total_cells - missing_cells) / total_cells) * 100, 2)
        }
    
    def analyze_missing_values(self):
        """Analyze missing values in the dataset."""
        missing_data = self.data.isnull().sum()
        missing_data = missing_data[missing_data > 0]
        
        if not missing_data.empty:
            missing_df = pd.DataFrame({
                'Missing_Count': missing_data,
                'Missing_Percentage': (missing_data / len(self.data)) * 100
            })
            return missing_df.sort_values('Missing_Count', ascending=False)
        return pd.DataFrame()
    
    def detect_outliers(self, method='iqr'):
        """Detect outliers in numerical columns."""
        numerical_cols = self.data.select_dtypes(include=[np.number]).columns
        outliers_info = {}
        
        for col in numerical_cols:
            if method == 'iqr':
                Q1 = self.data[col].quantile(0.25)
                Q3 = self.data[col].quantile(0.75)
                IQR = Q3 - Q1
                lower_bound = Q1 - 1.5 * IQR
                upper_bound = Q3 + 1.5 * IQR
                outliers = self.data[(self.data[col] < lower_bound) | (self.data[col] > upper_bound)][col]
            elif method == 'zscore':
                z_scores = np.abs(stats.zscore(self.data[col].dropna()))
                outliers = self.data[col][z_scores > 3]
            
            if len(outliers) > 0:
                outliers_info[col] = outliers
        
        return outliers_info
    
    def calculate_correlation_matrix(self):
        """Calculate correlation matrix for numerical columns."""
        numerical_data = self.data.select_dtypes(include=[np.number])
        if len(numerical_data.columns) > 1:
            return numerical_data.corr()
        return pd.DataFrame()
    
    def find_strong_correlations(self, threshold=0.7):
        """Find pairs of columns with strong correlations."""
        corr_matrix = self.calculate_correlation_matrix()
        if corr_matrix.empty:
            return pd.DataFrame()
        
        # Get upper triangle of correlation matrix
        upper_triangle = corr_matrix.where(
            np.triu(np.ones(corr_matrix.shape), k=1).astype(bool)
        )
        
        # Find strong correlations
        strong_corr = []
        for col in upper_triangle.columns:
            for idx in upper_triangle.index:
                value = upper_triangle.loc[idx, col]
                if not pd.isna(value) and abs(value) >= threshold:
                    strong_corr.append({
                        'Variable 1': idx,
                        'Variable 2': col,
                        'Correlation': round(value, 3),
                        'Strength': 'Strong Positive' if value > 0 else 'Strong Negative'
                    })
        
        return pd.DataFrame(strong_corr).sort_values('Correlation', key=abs, ascending=False)
    
    def analyze_categorical_distribution(self, column):
        """Analyze distribution of a categorical column."""
        if column in self.data.columns:
            value_counts = self.data[column].value_counts()
            return {
                'value_counts': value_counts,
                'unique_count': self.data[column].nunique(),
                'most_frequent': value_counts.index[0],
                'frequency_percentage': (value_counts / len(self.data)) * 100
            }
        return None
    
    def detect_data_patterns(self):
        """Detect common patterns in the data."""
        patterns = {}
        
        # Check for time series data
        datetime_cols = self.data.select_dtypes(include=['datetime64']).columns
        if len(datetime_cols) > 0:
            patterns['time_series'] = list(datetime_cols)
        
        # Check for potential ID columns
        id_columns = []
        for col in self.data.columns:
            if (self.data[col].nunique() == len(self.data) or 
                col.lower() in ['id', 'index', 'key'] or
                'id' in col.lower()):
                id_columns.append(col)
        if id_columns:
            patterns['id_columns'] = id_columns
        
        # Check for potential categorical columns with low cardinality
        categorical_candidates = []
        for col in self.data.select_dtypes(include=[np.number]).columns:
            if self.data[col].nunique() <= 10 and self.data[col].nunique() > 1:
                categorical_candidates.append(col)
        if categorical_candidates:
            patterns['potential_categorical'] = categorical_candidates
        
        return patterns
