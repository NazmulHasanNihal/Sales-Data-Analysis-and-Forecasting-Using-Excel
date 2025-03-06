# Sales Data Analysis and Forecasting Using Excel

This project demonstrates the process of sales data analysis and forecasting using Microsoft Excel. The goal of this project is to analyze historical sales data, identify trends and seasonal patterns, and apply time series forecasting techniques to predict future sales.

## Project Overview

**Objective:** 

Analyzing historical sales data, identify trends, seasonality, and key drivers of sales, and apply forecasting techniques to predict future sales. The goal is to use the data to provide actionable insights and optimize business strategies.

## Data Collection and Preparation

>
> * **Importing Data**
>   * Loading the sales dataset CSV into Excel.
>     * Ensuring proper data formatting (date, numbers, text).
> * **Data Cleaning**
>   * Removing any duplicates and handle missing values (fill or remove).
>     * Converting columns to appropriate data types (dates, numbers).
> * **Data Transformation**
>   * Creating new columns:
>     * Year, Month, Day of Week (derived from Date column).
>     * Sales per Transaction = Total / Quantity.

**The data is ready for further analysis or forecasting. We have cleaned the dataset, transformed it into a more useful format, and added calculated columns for deeper insights.**

## Time Series Analysis

> * **Trend Analysis**
>   * Plot sales data (Total) over time to identify trends.
>   * Applying linear regression to find an overall trend in the data.
> * **Seasonality Analysis**
>   * Using Pivot Tables to break down sales by Month.
>   * Detecting seasonal patterns.
> * **Moving Averages and Smoothing**
>   * Applying Simple Moving Average (SMA) and Exponential Moving Average (EMA) to smooth sales data.
>   * Using Exponential Smoothing (FORECAST.ETS) for trend and seasonality adjustments.

![Trend Analysis](img/Trend_Analysis.png)

**Summary:**

The chart shows a significant spike in sales amount and quantity in April, followed by steady growth in quantity sold but a flattening of sales revenue in the following months.

**Key Insights:**

> * April spike likely due to promotions or seasonality, driving both sales quantity and revenue up.
> * **Sales quantity** is growing steadily, but revenue growth has plateaued post-April, suggesting potential pricing or margin issues.
> * Focus on replicating April’s success and enhancing revenue per unit by adjusting pricing or product strategies.

![Seasonality Analysis By Quarters](img/Seasonality_Analysis_By_Quarters.png)

**Summary:**

The chart shows the sales amount across four quarters. Qtr1 has the lowest sales, while Qtr2, Qtr3, and Qtr4 show similar and higher sales figures.

**Key Insights:**

> * Qtr1 has significantly lower sales, suggesting a potential off-season or lower demand in the first quarter.
> * Qtr2, Qtr3, and Qtr4 have similar, higher sales, indicating stronger seasonal demand or consistent performance across these quarters.
> * Strategies should focus on boosting sales during Qtr1, leveraging insights from the stronger quarters.

![Exponential Moving Average](img/Exponential_Moving_Average.png)

**Summary:**

The chart shows Sales Amount over the months with an Exponential Moving Average (EMA) trendline. The EMA highlights the overall upward trend in sales, particularly after April.

**Key Insights:**

> * Sharp Increase in Sales: A significant spike in sales occurs in April, which is reflected in both the actual sales and the smoothed EMA.
> * EMA Trendline: The EMA line shows a steady upward trajectory, indicating a general increase in sales over the months.
> * Post-April Growth: After April, sales appear to maintain higher levels, suggesting sustained growth throughout the rest of the year.

## Key Driver Analysis

> * **Correlation Analysis**
>   * Performing correlation analysis to identify significant drivers of sales (`Unit price`, `Quantity sold`, `Product line`).
>   * Using scatter plots and correlation matrices to visualize relationships.
> * **Regression Analysis**
>   * Performing multiple regression analysis to model the relationship between sales and key variables (`Unit price`, `Quantity`, `Product line`, `Customer type`).
>   * Using Stepwise Regression for automatic selection of significant variables.
> * **Elasticity Analysis**
>   * Performing price elasticity analysis to measure how price changes affect sales.

### Correlation Analysis

![Sales vs Product Line](img/Sales_vs_Product_Line.png)
![Sales vs Quantity](img/Sales_vs_Quantity.png)
![Sales vs Unit Price](img/Sales_vs_Unit_Price.png)

**Summary:**

`Sales vs Product Line:` No strong relationship between sales and product line, as shown by the scattered data points.
`Sales vs Quantity:` A positive correlation exists, with sales increasing as quantity sold rises.
`Sales vs Unit Price:` There is a noticeable increase in sales as unit price increases, indicating that higher-priced items tend to generate higher sales.

**Key Insights:**

`Sales and Product Line:` Sales do not significantly vary by product line, suggesting other factors (like price or quantity) may drive sales.
`Quantity and Sales:` The strong positive relationship between sales and quantity indicates that increasing the quantity sold directly boosts total sales.
`Price and Sales:` The positive correlation between sales and unit price suggests that higher-priced products contribute to higher sales amounts.

### Regression Analysis

![Regression Analysis](img/Regression_Analysis.png)

**Summary:**

>
> * R-squared of `0.8855` indicates that `88.55%` of the variability in the dependent variable is explained by the model, suggesting a strong fit.
> * Adjusted R-squared `(0.8855)` is similar to R-squared, indicating no overfitting.
> * The F-statistic is `12879.32`, and Significance F is `0`, indicating the model is statistically significant.
> * The coefficients show the influence of each variable:
>   * `Intercept: -315.09`.
>   * `Unit price (74.69):` 
>     * Strong positive effect on the dependent variable.
>   * `Quantity (7)`: 
>     * Strong positive effect.
>   * `Product Line (1)`: 
>     * No significant effect on the dependent variable, as its p-value is high `(0.905)`.

**Key Insights:**
>
> * `Strong Model Fit:` The model explains a large portion of the variance in the dependent variable `(88.55%)`.
> * `Statistical Significance:` The regression model is statistically significant, confirming its validity.
> * `Key Drivers:` Unit price and Quantity significantly affect the dependent variable, while Product Line has a minimal effect.
