# Purpose of the program.
Saving all SolidWorks models from the assembly or parts in STL e.g. for 3D printing..

# Program capabilities.
1. Part file. Save all bodies from all configurations.
2. Assembly file. Opens each unique part in an assembly and processes it as a part file.
3. Assembly file containing assemblies. Not tested but provided :)


# Usage
1. Open the file with your project (assembly or part) in SolidWorks.
2. Run SolidWorksSaveSTL.exe --Out path
3. All solid bodies from SolidBodyFolder will be recorded from all parts of all configurations in them.

The file name is taken from the name of the solid body.

# Parameters
--Out Path. Path for saving files.

