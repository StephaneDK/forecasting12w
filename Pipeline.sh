#!/bin/sh

source "C:\Users\steph\miniconda3\etc\profile.d\conda.sh"
#source "C:\Users\snichanian\miniconda3\etc\profile.d\conda.sh"
conda activate forecast_env

#Forecasting paths
path_local='C:\Users\steph\Documents\DK\Work\Forecasting book sales and inventory\Pipeline\Code\'
path_DK='C:\Users\snichanian\Documents\DK\Work\Forecasting book sales and inventory\Pipeline\Code\'

#Forecasting Country
country="us"

#Change this line for directory
path_code=$path_local

export path_code
export country

cd "$path_code"

outputString=$(python "fetch_data_snfk_dly.py" 2>qtemp);

stderrString=$(<qtemp);
echo "$outputString"

if [[ "$stderrString" == "ValueError1" ]]; then
    printf "\nExecution halted"
    exit
fi

if [[ $country == "uk" ]]; then
    Rscript "Slowing trending titles UK.R" 
    Rscript "Full Year Forecast UK.R" 
     
elif [[ $country == "us" ]]; then
    Rscript "Slowing trending titles US.R"
    Rscript "Full Year Forecast US.R"
    Rscript "Bookscan Forecast US.R"
     
fi

python "write_data_snfk.py" $country

conda deactivate








