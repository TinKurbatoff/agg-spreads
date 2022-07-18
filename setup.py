import setuptools

setuptools.setup(
    name="agg_spreads",
    version="0.1.0",
    author="Constantine K",
    description="One more gspread wraper to work with Google Sheets",
    packages=["agg_spreads"],
    install_requires=[
        'google-api-python-client',
        'pandas',
    ])
