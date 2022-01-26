#!/bin/sh

country='us'

outputString=$(python "C:\Users\Stephane\Documents\DK\Work\Forecasting book sales and inventory\Pipeline\Code\Fetch_data.py" $country 2>qtemp);
stderrString=$(<qtemp);
echo "$outputString"


if [[ "$stderrString" == "ValueError1" ]]; then
    printf "\nExecution halted"
    exit
fi

if [[ $country == "uk" ]]; then
    Rscript "C:\Users\Stephane\Documents\DK\Work\Forecasting book sales and inventory\Pipeline\Code\Slowing trending titles UK.R" 
    Rscript "C:\Users\Stephane\Documents\DK\Work\Forecasting book sales and inventory\Pipeline\Code\Full Year Forecast UK.R" 
    Rscript "C:\Users\Stephane\Documents\DK\Work\Forecasting book sales and inventory\Pipeline\Code\Blacklist UK.R" 
elif [[ $country == "us" ]]; then
    Rscript "C:\Users\Stephane\Documents\DK\Work\Forecasting book sales and inventory\Pipeline\Code\Slowing trending titles US.R"
    Rscript "C:\Users\Stephane\Documents\DK\Work\Forecasting book sales and inventory\Pipeline\Code\Full Year Forecast US.R"
    Rscript "C:\Users\Stephane\Documents\DK\Work\Forecasting book sales and inventory\Pipeline\Code\Blacklist US.R" 
     
     
fi

python "C:\Users\Stephane\Documents\DK\Work\Forecasting book sales and inventory\Pipeline\Code\Write_data.py" $country

python "C:\Users\Stephane\Documents\DK\Work\Forecasting book sales and inventory\Pipeline\Code\Delete_Duplicate.py" 


#python "C:\Users\Steph\Documents\DK\Work\Forecasting book sales and inventory\Pipeline\Code\Update_View.py"

#Rscript "C:\Users\Steph\Documents\DK\Work\Forecasting book sales and inventory\Pipeline\Code\Update Dashboard.R"








