# Stripe Reporter

The Stripe Reporter is a script that generates an overview of Stripe payouts, including details of what was sold and the associated fees. It provides an easy-to-use solution for generating an Excel file that can be used for bookkeeping purposes.

## Features

- Retrieves Stripe payouts and extracts transaction details.
- Handles different types of transactions, including SudoSOS top-ups and product sales.
- Generates an overview report for each Stripe payout, including product details, transaction amounts, fees, and more.
- Saves the generated report to an Excel file for easy record-keeping.

## Requirements

- Python 3.x
- Stripe API credentials (API key)

## Installation

1. Clone the repository:  
`git clone https://github.com/GEWIS/stripe-reporter.git`

2. Install the required dependencies:  
`pip install -r requirements.txt`

3. Copy `.env.example` to `.env` and set up your Stripe API key.

## Usage

`python stripe-reporter.py [-h] [-s N] [-j JSON] [-c] [-x] [-n NAME]`

### Flags

- `-h, --help`: Show the help message and exit.
- `-s N, --poll-stripe N`: Select id from the latest N StripePayout.
- `-j JSON, --json JSON`: Read the report data from a JSON file provided by the flag.
- `-c, --print-json`: Print the JSON result from `get_report_data` to stdout.
- `-x, --save-to-excel`: Save the report data to an Excel file.
- `-n NAME, --output-name NAME`: Specify the output name for the Excel file. If not provided, the default name will be `"Report_{PayoutId}.xlsx"`.

## Notes

- SudoSOS top-ups are treated as direct charges without specific product information in Stripe. This script handles them accordingly.
- Payment links with attached products are used to sell other products, and their details are included in the generated overview.
- **This scripts requires an API call for each transaction in the payout. This means it is not suited for large payouts. (Reference 100 transactions take about 7 seconds to report.)**