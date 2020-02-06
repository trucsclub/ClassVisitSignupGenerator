# Class Visit Signup Sheet Generator

This application will automatically generate a class visit signup sheet in Excel format. All you have to do is supply it with a JSON file of all the classes.

How do you get a JSON file with all the classes, you ask? Simple. Well, not as simple as if I had bothered to add login. Or maybe I just didn't want to share that code ;)

At this time, you need to

1. Sign in to MyTRU
2. Go to "Course Planning and Registration", then "Browse classes". Then, select the current semester.
3. First, add "COMP" into the Subject. Then, open the Network inspector. Look for the JSON XHR to searchResults?txt_subject=..., and open it in a new tab. Edit "pageMaxSize" to something like 150. It needs to be bigger than the number of classes offered that semester.
4. Save in the same folder as this program as "classes.json"

Simple, right? By doing this, you will officially be a Banner Hacker.

## Downloading

See [releases](https://github.com/trucsclub/ClassVisitSignupGenerator/releases).

## Building

You will need the .NET 3.1 SDK (or higher).

To build from the command line, run ``dotnet build``. This will generate a debug build.

To run what you just built, cd into the project directory and run ``dotnet run``.

To publish, specifically in a single-binary release, run ``dotnet publish -r {RID} -c Release /p:PublishSingleFile=true /p:PublishTrimmed=true``. Replace ``{RID}`` with the [runtime identifier](https://docs.microsoft.com/en-us/dotnet/core/rid-catalog) of the platform you are targetting. For example, for 64-bit Windows use ``win-x64``. 
