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

The file name is generated from:
1. The name of the component.
2. Configuration name, if more than 1
3. Name of the body, if more than 1.

$"{partName} {configName} {bodyName}"

# Parameters
--Out Path. Path for saving files.

# Problems
The project contains files from SolidWorks 2021. I have not tried it with other versions. In theory, it should work with any versions>=2021 but most likely even younger. In case of problems, you can try to replace the files from the folder with the installed program C:\Program Files\SOLIDWORKS Corp\SOLIDWORKS\api\redist
