# What is this?
<br>

This is basically a python based subreddit analytics bot which scrapes data of the subreddit you enter

<br>

# Requirements
## Pip and Git
<br>

You should have preinstalled python, pip, and git. You can install python and pip together in the latest version of [python](https://www.python.org/downloads/). As for git, you can download it from [here](https://git-scm.com/). If you are on Mac, you can simply run `git --version` in terminal. If you are on Linux, you can simply run `sudo apt install git` in the terminal

<br>

## A Reddit Application
<br>

Okay so first of all, log in to your reddit account. Then go to reddit.com/prefs/apps. Scroll down to the bottom and click "are you a developer? create an app..."

![](https://i.imgur.com/bGY3l9r.png)

Then select script

![](https://i.imgur.com/C44zSwM.png)

Then enter any random information in "name", "description", "about url", and "redirect url". Then press "create app"

![](https://i.imgur.com/ePB3BeE.png)

Now your screen should look something like this:

![](https://i.imgur.com/qPnrS56.png)

Now the text in the top red box is called client ID, and the text bottom red box is called Client Secret. Copy them somewhere since these two things will be neccessary in the installation of the program

<br>

# Installation
<br>

Okay so now, enter the following commands in your command prompt/terminal:

```
git clone https://github.com/stupid-melon/subreddit-analytics-bot
cd subreddit-analytics-bot
pip install -r requirements.txt
```
(Replace `pip` with `pip3` if you are on linux or mac)

<br>

Now that we have installed everything, its now time to set the configurations. In the same command prompt/terminal window, enter:

<br>

(If you are on windows)
```
start config.txt
```

(If you are on mac/linux)
```
open config.txt
```

<br>

This should open up a file and it should look something like this:

![](https://i.imgur.com/Nf7qtCH.png)

Now in here, enter your client ID and client secret in the double quotes like this:

![](https://i.imgur.com/WF9dwaY.png)

Now save the file and close both, the command prompt/terminal window as well as the file

<br>

# Usage
<br>

Now whenever you want to use this program, just open up a command prompt/terminal window and enter:

```
cd subreddit-analytics-bot
python main.py
```

(Replace `python` with `python3` if you are using mac or linux)

<br>

If you ever want to change the configuration, enter in command prompt/termnial:

```
cd subreddit-analytics-bot
start config.txt
```

(Replace `start` with `open` if you are on macOS)

<br>

In the config file, set `ask` to `"False"` if you dont want the computer program to always ask which subreddit you want to analyse. Instead, it would analyse the subreddit given in `subreddit`

<br>

The output is saved in a .xlsx file in the output folder whose location is specified once the program is finished. The .xlsx file contains 4 sheets namely:

<br>

General Info

![](https://i.imgur.com/tUnRt2H.png)

Top Posts

![](https://i.imgur.com/eaPCWVQ.png)

Top Posters

![](https://i.imgur.com/BGAV2mp.png)

Top Words

![](https://i.imgur.com/9RQNJrx.png)
