### To build

build with

> nexe.cmd .\index.js -o .\build\combine.exe

### to run

- set the load and save paths, Load is the path to all the subdir to searh though, save is where to create the combined file
- set unstyled, styled file names for the filename of the middle and final files created (use same name to have it not keep the unstyled middle file)
- instead of settings file you can use env vars: LOAD_PATH, SAVE_PATH, UNSTYLE_FILENAME, STYLE_FILENAME
- the build will include the settings file or env vars when you build, to change paths or filenames you must build again
- have all the parts to combine in subfolders of your load path
- you can run with npm 

