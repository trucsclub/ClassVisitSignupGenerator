# Class Visit Signup Sheet Generator

This application will automatically generate a class visit signup sheet in Excel format. All you have to do is supply it with a JSON file of all the classes.

How do you get a JSON file with all the classes, you ask? Simple. Well, not as simple as if I had bothered to add login. Or maybe I just didn't want to share that code ;)

At this time, you need to

1. Sign in to MyTRU
2. Go to "Course Planning and Registration", then "Browse classes". Then, select the current semester.
3. First, add "COMP" into the Subject. Then, open the Network inspector. Look for the JSON XHR to searchResults?txt_subject=..., and open it in a new tab. Edit "pageMaxSize" to something like 150. It needs to be bigger than the number of classes offered that semester.
4. Save in the same folder as this program as "classes.json"

Simple, right? By doing this, you will officially be a Banner Hacker.