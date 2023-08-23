<!-- Improved compatibility of back to top link: See: https://github.com/othneildrew/Best-README-Template/pull/73 -->
<a name="readme-top"></a>
<!--
*** Thanks for checking out the Best-README-Template. If you have a suggestion
*** that would make this better, please fork the repo and create a pull request
*** or simply open an issue with the tag "enhancement".
*** Don't forget to give the project a star!
*** Thanks again! Now go create something AMAZING! :D
-->



<!-- PROJECT SHIELDS -->
<!--
*** I'm using markdown "reference style" links for readability.
*** Reference links are enclosed in brackets [ ] instead of parentheses ( ).
*** See the bottom of this document for the declaration of the reference variables
*** for contributors-url, forks-url, etc. This is an optional, concise syntax you may use.
*** https://www.markdownguide.org/basic-syntax/#reference-style-links
-->


<!-- TABLE OF CONTENTS -->
<details>
  <summary>Table of Contents</summary>
  <ol>
    <li>
      <a href="#about-the-project">About The Project</a>
    </li>
    <li>
      <a href="#disclaimer">Disclaimer</a>
    </li>
    <li>
      <a href="#getting-started">Getting Started</a>
      <ul>
        <li><a href="#prerequisites">Prerequisites</a></li>
        <li><a href="#Requirements">Requirements</a></li>
      </ul>
    </li>
    <li><a href="#roadmap">Roadmap</a></li>
    <li><a href="#contact">Contact</a></li>
  </ol>
</details>



<!-- ABOUT THE PROJECT -->
## About The Project

This repository is for the forecast over the next 12 weeks of DK's book sales. It is comprised of a multitude of scripts, in R Python and Shell.

The python files are used as an interface to snowflake, to download sales data, inventory data, reprint data, AA and print status. 
We also use the Python files to write the resulting forecasts into the database in the following tables:

* PRH_GLOBAL_DK_SANDBOX.PUBLIC.FORECASTING_TEMP
* PRH_GLOBAL_DK_SANDBOX.PUBLIC.FORECASTING_STATISTICS_TEMP

The R files are used to create the forecasting models and execute the calculations. Once the forecasts are completed, the R file will save the output in a xlsx file which is then shared with our inventory teams.

Finally, the Shell script is used to sequentially call the various forecasting scripts.

A central part of this repository is reproducibility. This repository is designed to assist D&A team members to run the forecasting pipeline from their DK work laptops smoothly.  
For this purpose, it is suggested that a virtual environment is created with all the necessary libraries and packages provided in the requirements.txt and renv.lock files.

<p align="right">(<a href="#readme-top">back to top</a>)</p>

<!-- DISCLAIMER -->
## Disclaimer
The forecasting pipeline is currently locally run, meaning that it is only possible to run the code from a DK laptop (rather than the cloud). This means that all of the directories specified in the various scripts must be change to your own directories.

<p align="right">(<a href="#readme-top">back to top</a>)</p>

<!-- GETTING STARTED -->
## Getting Started

This section will explain the steps to take to be able to run the forecasting pipeline in its entirety.

### Prerequisites

In order to run the forecasting pipeline, the following coding languages must be installed on your computer:

* R (see following link [https://cran.r-project.org/bin/](https://cran.r-project.org/bin/) )
* Python (see following link [https://www.python.org/downloads/](https://www.python.org/downloads/) )

The repository can be cloned using the following command from git:
   ```sh
   git clone https://github.com/StephaneDK/forecasting12w-remote.git
   ```


### Requirements

It is strongly suggested to use a virtual environment for this project. Given that there are two coding languages, there are also two virtual environments to be used.

**Create a Python Conda Environment**
   ```sh
   conda create -n ENV_NAME  python=3.10
   ```

**Activate Python Conda Environment**
   ```sh
   conda activate ENV_NAME 
   ```

**Installing required libraries from requirement.txt**
   ```sh
   pip install -r requirements.txt 
   ```


**Create a R Environment**
   ```sh
   renv::init()
   ```

**Activate R Environment**
   ```sh
   renv::activate() 
   ```

**Install required packages from renv.lock**
   ```sh
   renv::restore() 
   ```

The code is set up so that the R virtual environment is activated and deactivated in the Rscript files directly, so this step doesn't need to be run explicitly.

The required libraries to be installed can be found in the requirements.txt and renv.lock files.


<p align="right">(<a href="#readme-top">back to top</a>)</p>



<!-- ROADMAP -->
## Roadmap

- [x] Creation of forecasting models
- [x] Automation of inventory data
- [ ] Full Automation of forecasting pipeline
- [ ] Migration of forecasting pipeline in AWS cloud server
- [ ] Explore further models: front list and NYP

<p align="right">(<a href="#readme-top">back to top</a>)</p>


<!-- CONTACT -->
## Contact

Stephane Nichanian -  stephane.nichanian@uk.dk.com

Project Link: [https://github.com/StephaneDK/forecasting12w-remote](https://github.com/StephaneDK/forecasting12w-remote)

<p align="right">(<a href="#readme-top">back to top</a>)</p>

