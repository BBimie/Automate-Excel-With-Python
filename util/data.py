import pandas as pd

def sales_data() -> pd.DataFrame:
    """Creates a Pandas DataFrame containing sales data for a fictional company.

    The DataFrame has four columns: date, item, cost_price, and sale_price. The 'date' column
    contains dates in the format 'mm/dd/yyyy', while the 'item', 'cost_price', and 'sale_price'
    columns contain information about the products sold, their cost prices, and their sale prices,
    respectively. The data is manually typed and serves as a sample dataset for demonstration purposes.

    Returns:
    -------
    Pandas DataFrame
        A DataFrame containing the sales data for the fictional company.
    """

    data = {
        'Date': [
            '01/01/2023', '01/01/2023', '01/01/2023', '01/02/2023', 
            '01/02/2023', '01/03/2023', '01/04/2023', '01/04/2023', 
            '01/05/2023', '01/05/2023', '01/06/2023', '01/06/2023', 
        ],
        'Items': [
            'Bookcase', 'Chair', 'Table', 'Bookcase', 'Chair', 'Table', 
            'Shelf', 'Shelf',  'Chair', 'Table', 'Chair', 'Table',
        ],
        'Cost Price': [
            15000, 12500, 10000, 4000, 2500, 13500, 12000, 15000, 
            12500, 10000, 4000, 2500,
        ],
        'Sale Price': [
            18000, 15000, 17000, 6500, 4550, 18570, 14650, 18500, 
            15450, 17800, 6900, 4950,
        ]
    }

    return pd.DataFrame(data)
