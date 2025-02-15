## setup
- install clasp : `npm install -g @google/clasp` 
- create empty dir `src`
- `clasp login`
- npm init
- grab your appscript project id
- `clasp clone <id> --rootDir src`
- `clasp push -w` to watch for changes

## development
- `clasp push -w` on local to push changes. `-w` is for continuous watch
- `clasp pull` for pull changes from the remote appscript project.
- debug on browser with a debugger sidebar.
