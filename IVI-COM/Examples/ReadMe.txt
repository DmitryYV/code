AgilentN67xx Programming Examples
-------------------------------------

Most examples are simple console applications that demonstrate basic use 
of the driver in a variety of application development environments and 
programming languages.  All the example programs do the following:

1. Create an instance of the driver.
2. Read driver Identity properties.
3. Initialize the driver.
4. Read instrument Identity properties.
5. Check the instrument error queue.
6. Close the driver.

Notes:
-----
Some examples may perform additional driver specific tasks.

Examples should compile and run as is, assuming your IVI directory is in
the default location of:  C:\Program Files\IVI\

The VS6, C++ example may be converted and run as is in VS.NET if desired.

The Agilent VEE Pro 7.5 adds significant new features for using IVI-COM drivers.
Two example programs are provided, one for use with versions prior to 7.5 and
another for version 7.5 and up which uses the new features.  The older example
is fully compatible with VEE 7.5 or greater.

You may use these examples as a starting point for creating programs for
your driver.  For additional info on using IVI-COM drivers in a variety
of development environments, see the Getting Started section of the driver
Help file.
