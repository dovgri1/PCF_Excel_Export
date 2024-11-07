# Dataverse PCF Control Library

Welcome to the **PCF Excel Export project**! This repository contains custom Power Platform Component Framework (PCF) control purposed to create custom export to excel logic for you dataverse environment.

## Installation

### Prerequisites

- Node.js (latest stable version) https://nodejs.org/en/download/source-code
- Power Platform CLI (`pac`) https://aka.ms/PowerAppsCLI
- Fiddler Classic Client https://www.telerik.com/download/fiddler
- Dataverse or Power Apps environment
- PCF Project Template

### Installation Steps

1. Clone this repository
2. Run 'npm install' in your terminal
3. To create authentication profile: pac auth create -env https://yourenvironment.crm4.dynamics.com/
4. To see all crated profiles: pac auth list
5. To select your auth profile: pac auth select -i [index of your create profile]
6. To push your component to selected environment: pac pcf push --solution-unique-name YourSolutionNameHere

### Fiddler setup 

When you component is deployed to environment, you can start configure your fiddler client. In Auto Responder section 2 rules need to be added for each environment:
1. Top input: REGEX:(.*?)(\/css)?(\/|cc_)[your component namespace].[your component name].(?'path') Bottom Input: *header:Cache-Control=no-cache, no-store, must-revalidate
2. Top Input: REGEX:(.*?)(\/css)?(\/|cc_)[your component namespace].[your component name].(?'path') Bottom Input: [Path to the folder where budle js is built]\$2\${path} example: C:\repos\Macaw\PCF_Excel_Export_Component\out\controls\pcfexcelexport\$2\${path}
Onle rules are created don't forget to click check box named 'Enable rules'
