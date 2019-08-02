# Overview

The kadviz is a command line utility is used to enumerate an Active Directory (AD) groups and present the group, its users and child or parent groups (forward and reverse lookup) of the original group in an easy to use format to facilitate AD re-organisation.

Currently, two output modes are supported: as comma separated value file (CSV) and as a directed graph using GraphViz graphing engine.

![Sample output](https://github.com/truekonrads/kadviz/raw/master/gallery/example.png)

# Installation

kadviz uses GraphViz to draw graphs. The GraphViz is a free, open source graphing package which can be [downloaded from the official site] (http://www.graphviz.org/Download..php)

You will need to install it to be able to draw graphs.

```
usage: kadviz [-h] [--debug] [-o FILE] [--no-recursive]
              [-f {csv,pretty-print,graph-dot}] [-s SOURCE]
              [-d {reverse,forward} | -A] [--fetch-users | --no-fetch-users]
              [--no-dn] [--graphviz-dot DOTFILE] [--graph-users]
              [--credentials CREDENTIALS | --credfile CREDFILE]
              OBJECT

Retrieve and expand AD groups. If no credentials are supplied current
credentials of logged on user will be used

positional arguments:
  OBJECT                Object (group or user) to expand. You can specify both
                        DN and NT names

optional arguments:
  -h, --help            show this help message and exit
  --debug               Enable debugging output
  -o FILE, --output FILE
                        Output file, defaults to stdout
  --no-recursive        Do not expand child groups
  -f {csv,pretty-print,graph-dot}, --format {csv,pretty-print,graph-dot}
                        Output format
  -s SOURCE, --data-source SOURCE
                        Specify data source. Defaults to global catalogue GC:
  -d {reverse,forward}, --direction {reverse,forward}
                        Lookup direction. Forward looks up the members of the
                        group (only for groups), reverse - groups that the
                        specified object is a member of. The default for
                        groups is forward, while for users - reverse
  -A, --dump-all        Dump every group

User expansion:
  Options related to user expansion. By default, if input object is group,
  then users will be graphed, if user, then not

  --fetch-users         Always force users
  --no-fetch-users      Never fetch users

csv:
  Arguments relating to CSV output

  --no-dn               Do not print DNs in CSV

graphing:
  Options related to graph generation

  --graphviz-dot DOTFILE
                        Path to GraphViz dot.exe, best guess: NOT FOUND
  --graph-users         Expand graph users, don't summarise

creds:
  credential management

  --credentials CREDENTIALS
                        Credentials separated by column like bp1\user:password
  --credfile CREDFILE   Credential file which contains one line with rot13 +
                        base64 encoded credentials like bp1\user:password

Supported output formats are: * csv - dump to a CSV file * pretty-print - dump
data structures in indented fashion (for debugging) * graph-dot - output a
.dot file which is agraph defintion as understood by GraphViz. By Konrads
Klints <konrads.klints@kpmg.co.uk>.
```
