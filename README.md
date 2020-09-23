# VBA-Challenge
This script filters through the data provided in each sheet and provides a single row value for each ticker.
Accompanying that ticker is a yearly change, percent change & total volume.
The yearly change is calculated by taking the first open value of a ticker that does not equal 0 and subtracting it by the last available close value of that ticker.
The percent change is calculated by dividing the yearly change by the first open value.
The total volume is a sum of all available volume for a ticker year.
