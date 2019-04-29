How To:


    1.Add ModifyInMemory.cs to project

    2.Activate MemoryPatching in entry method of application:

    LicenseHelper.ModifyInMemory.ActivateMemoryPatching();

    Notes:
    - The "SetLicense"-method is invoked automatically for all assemblies !
    - MemoryPatch bypasses 1024bit RSA protection

    3.Compile project with option "Allow unsafe code"
    (=> Project-Properties > Build: Check according option!)


Notes:

    MemoryPatching is valid as long as the licensing does not change
    (If I remember right in the last years nothing has changed)

    ModifyInMemory.cs must be excluded from obfuscation

    No support for Silverlight because SL's mscorlib.dll does not include all methods/properties required for MemoryPatching 
=================
www.downloadly.ir