# SA Reporting tool

## Setup dev environment

Setting up on VSCode:
- install node
  - `sudo dnf install nodejs`
- install autocompletion
  - `npm install --save @types/google-apps-script`
- Install clasp
  - `npm install  @google/clasp`


- Get Project Id. (from apps script settings page) 
- [Enable](https://script.google.com/home/usersettings) Apps Script API
- Login 
  - `clasp login`
- Push changes
  - `clasp push`