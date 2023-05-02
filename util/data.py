import pandas as pd

def sales_data():
    ''' The function creates a simple dataframe with manually typed data, the dataframe will have four
        columnns showing a company's sales information.'''

    data = {'date': ['01/01/2023', '01/02/2023', '01/03/2023', '01/04/2023', '01/01/2023', '01/02/2023', '01/03/2023',
                    '01/05/2023', '01/06/2023', '01/07/2023', '01/05/2023', '01/06/2023', '01/07/2023', '01/01/2023'],
            'Items': ['Bookcase', 'Chair', 'Table', 'Art', 'Frame', 'Bookcase', 'Chair', 
                    'Bookcase', 'Chair', 'Table', 'Art', 'Frame', 'Bookcase', 'Chair', ],
            'Cost Price': [15000, 12500, 10000, 4000, 2500, 13500, 12000,
                        15000, 12500, 10000, 4000, 2500, 13500, 12000],
            'Sale Price': [18000, 15000, 17000, 6500, 4550, 18570, 14650,
                        18500, 15450, 17800, 6900, 4950, 19500, 14350] }

    df = pd.DataFrame(data)
    return df