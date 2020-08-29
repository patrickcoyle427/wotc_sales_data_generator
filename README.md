# wotc_sales_data_generator

Takes a CSV from lightspeed retail POS, pulls out the sales information that Wizards
Of the Coast needs, then writes them to an xlsx file that can be sent for reporting.

-Written by Patrick Coyle

# USAGE:

- On first run of this script, two directories will be created in the same directory as this
  file: to_parse and parsed_data

- Place Lightspeed Retail POS sales data into the to_parse folder, these files are .csv files
  that are created in sales by line in the reporting section of Lightspeed Retail POS

- WOTC wants 1 month of all Magic: The Gathering and Dungeons and Dragons Sales. To create
  the data that will be used for this script, In Lightspeed POS, go to Reports, then Sales
  Lines. Set the date range to be the month you want sales from. For Magic: The Gathering
  search the "MTG" tag, and exclude "accessory". For Dungeons and Dragons, search the "dnd"
  tag and exclude "mini".
