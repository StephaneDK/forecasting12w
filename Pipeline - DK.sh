#!/bin/sh

country="us"


outputString=$(python "C:\Users\snichanian\Documents\DK\Work\Forecasting book sales and inventory\Pipeline\Code\fetch_data_snfk_dly.py" $country 2>qtemp);

stderrString=$(<qtemp);
echo "$outputString"

if [[ "$stderrString" == "ValueError1" ]]; then
    printf "\nExecution halted"
    exit
fi

if [[ $country == "uk" ]]; then
    Rscript "C:\Users\snichanian\Documents\DK\Work\Forecasting book sales and inventory\Pipeline\Code\Slowing trending titles UK.R" 
    Rscript "C:\Users\snichanian\Documents\DK\Work\Forecasting book sales and inventory\Pipeline\Code\Full Year Forecast UK.R" 
     
elif [[ $country == "us" ]]; then
    Rscript "C:\Users\snichanian\Documents\DK\Work\Forecasting book sales and inventory\Pipeline\Code\Slowing trending titles US.R"
    Rscript "C:\Users\snichanian\Documents\DK\Work\Forecasting book sales and inventory\Pipeline\Code\Full Year Forecast US.R"
    
     
     
fi

python "C:\Users\snichanian\Documents\DK\Work\Forecasting book sales and inventory\Pipeline\Code\Write_data.py" $country
python "C:\Users\snichanian\Documents\DK\Work\Forecasting book sales and inventory\Pipeline\Code\write_data_snfk.py" $country

Rscript "C:\Users\snichanian\Documents\DK\Work\Forecasting book sales and inventory\Pipeline\Code\Bookscan Forecast US.R"








