import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import os
import sys

def main():
        # Get the directory where the script is located
    if getattr(sys, 'frozen', False):
        # If the application is run as a bundle, the PyInstaller bootloader
        # extends the sys module by a flag frozen=True and sets the app 
        # path into variable _MEIPASS'.
        application_path = sys._MEIPASS
    else:
        application_path = os.path.dirname(os.path.abspath(__file__))
    
    # Construct the path to the Excel file
    excel_path = os.path.join(application_path, "BIGSheet.xlsx")
    # Read the Excel file once

    if not os.path.exists(excel_path):
        print(f"Excel file not found at {excel_path}")
        sys.exit(1)


    df = pd.read_excel(excel_path)

    # creates the basic arrays for the pay and payLabels
    payArray = []
    payLabelsArray = []

    # Extract the price column assuming it's in column 'C'
    price = df.iloc[:, 2] 

    # Extract the grid choice column assuming it's in column 'D'
    gridChoice = df.iloc[:, 3]
    gridChoice = gridChoice.iloc[0]

    # Extract the pay column assuming it's in column 'B'
    pay = df.iloc[:, 1]  

    # Extract the payLabels column assuming it's in column 'A'
    payLabels = df.iloc[:, 0]

    # Convert the extracted columns to arrays
    payArray = pay.values
    payLabelsArray = payLabels.values

    # Plot the price values
    WTP(price, payArray, payLabelsArray, gridChoice)

def WTP(price, payArray, payLabelsArray, gridChoice):

    # sets the gridlines to the payArray
    gridLines = payArray
    # finds the number of producers for the graph
    length = len(payArray) + 1

    # adds all neccaesary points to the payArray
    adjustedPayArray = []
    for i in payArray:
        adjustedPayArray.append(i)
        adjustedPayArray.append(i)
    adjustedPayArray.append(0)
    payArray = adjustedPayArray

    # Draw a horizontal line at the first price point
    plt.axhline(y=price.iloc[0], color='g', linestyle='-') 

    #creates all the points for the labels and adds them to the payLabelsArray
    adjustedPayArrayLabels = []
    for i in payLabelsArray:
        adjustedPayArrayLabels.append(i)
        adjustedPayArrayLabels.append(i)
    adjustedPayArrayLabels.pop(0)
    payLabelsArray = adjustedPayArrayLabels
    payLabelsArray.append("")
    payLabelsArray.append("")

    # generates the x values that the labels will be mapped to
    xValues = []
    for i in range(length):
        xValues.append(i)
        xValues.append(i)
    xValues.pop(0)

    # plots array into a line graph
    x = np.array(xValues)
    y = np.array(payArray)
    my_xticks = payLabelsArray
    plt.xticks(x, my_xticks)
    plt.plot(x, y, color="brown")
    
    # adds gridlines to the graph based on graph
    if gridChoice == 1:
        plt.grid(True)
        gridLines = list(gridLines)
        gridLines.pop(0)
        xAxis = 1
        for i in gridLines:
            y1 = np.array([i, 0])
            x1 = np.array([xAxis, xAxis])
            plt.plot(x1, y1, color="brown", linestyle="dashed")
            xAxis = xAxis + 1
        
    # sets 0,0 at the origin
    plt.xlim([0, max(x)+0.5])
    plt.ylim([0, max(y)+0.5])

    # labels the x and y axis
    plt.title("Willingness To Pay Model", color='black')
    plt.xlabel("Soil Producers", color='brown')
    plt.ylabel("Price", color='green')

    # shows the graph
    plt.show()

if __name__ == "__main__":
    main()