# Replication plot
# Author: Konrads Klints <konrads.klints@kpmg.co.uk>
import sys
import logging
import adsilib
import csv
import argparse
import win32com
import win32com.client as win32
import collections
import pprint
import pydot
import os,glob,tempfile,subprocess,re
import ldapaccess
import math



sys.setrecursionlimit(3000)
def _filter_node_badchars(n):
    assert type(n)==str or type(n)==unicode
    return re.sub("[^A-Za-z0-9_]","_",n)
    #return n.replace("&- ","_")

# def _makeTreeEntry(node,rel):
#     assert rel in "forward reverse none".split(" "),\
#            "Unknown relationship type between nodes %s" % rel
#     tree={'name':node['name'],'attrs':node,'users':[],'groups':[],'rel':rel}
#     return tree

def _find_graphviz_dot():
    candidates=glob.glob(os.path.join(os.environ['PROGRAMFILES'],"GraphViz*","bin","dot.exe"))
    if len(candidates):
        dotfile=candidates[0]
        if os.path.exists(dotfile):
            return dotfile
    return None

def getSiteLinks(source):
    attributes="cost cn replInterval siteList".split(" ")
    results=source.search('(objectClass=siteLink)',attributes)
    return results

def getSites(source):
    attributes="cn location".split(" ")
    results=source.search('(objectClass=site)',attributes)
    return results

def main():
    format_types="graph-dot graph".split(" ")
    #lookup_directions="reverse forward both".split(" ")
    lookup_directions="reverse forward".split(" ")
    parser = argparse.ArgumentParser(description="Retrieve and plot Active Directory sites and replication.\nIf no credentials are supplied current credentials of logged on user will be used",
                                     epilog=\
"""
Supported output formats are:
 * graph-dot - output a .dot file which is a graph defintion as understood by GraphViz.
 * graph     - build a pretty picture using the .dot file
 By Konrads Klints <konrads.klints@kpmg.co.uk>.
                                 """.strip())
    # parser.add_argument('groupName',metavar='OBJECT',type=str,nargs=1,
    #                     help='Object (group or user) to expand. You can specify both DN and NT names')    
    parser.add_argument('--debug',dest='debug', action='store_true', default=False,
                        help='Enable debugging output')
    parser.add_argument('-o','--output',dest='output', type=str,metavar="FILE",
                        help='Output file, defaults to stdout') 
    # parser.add_argument('--no-recursive',dest='recursive', action="store_false",default=True,
    #                     help='Do not expand child groups')   
    
    parser.add_argument('-f','--format',dest='format_type', type=lambda x: x.lower(),choices=format_types,
                        default=format_types[0],
                        help="Output format")
    parser.add_argument("-s","--data-source",
                        help="Specify data source. Defaults to global catalogue GC:",
                        type=str,default="GC:",dest='source')
#     dumpwhat=parser.add_mutually_exclusive_group()  
#     dumpwhat.add_argument("-d","--direction",
#             help="""Lookup direction. Forward looks up the members of the group (only for groups), 
#             reverse - groups that the specified object is a member of. 
#             The default for groups is forward, while for users - reverse""",
#             type=lambda x: x.lower(),choices=lookup_directions,
#             dest='direction')
#     dumpwhat.add_argument("-A","--dump-all",
#                           help="Dump every group",
#                           dest='dump_all',default=False,action="store_true")
#     ugg=parser.add_argument_group("User expansion",
#                                   """
# Options related to user expansion. By default, if input object is group, then users will be graphed, 
# if user, then not"""
#                                   ).add_mutually_exclusive_group()    
#     ugg.add_argument('--fetch-users',dest='always_fetch_users', action="store_true",    default=None ,                         
#                      help="Always force users")
#     ugg.add_argument('--no-fetch-users',dest='never_fetch_users', action="store_true", default=None,                           
#                      help="Never fetch users")    
    csvgroup=parser.add_argument_group('csv','Arguments relating to CSV output')
    csvgroup.add_argument('--no-dn',dest='csv_output_dn', action="store_false",default=True,                            
                          help="Do not print DNs in CSV")
    graphgroup=parser.add_argument_group("graphing","Options related to graph generation")
    dotfile=_find_graphviz_dot()
    if dotfile is None:
        dotfile="NOT FOUND"
    graphgroup.add_argument('--graphviz-dot',dest='dotfile', type=str,default=dotfile,                            
                            help="Path to GraphViz dot.exe, best guess: %s" % dotfile)
    # graphgroup.add_argument('--graph-users',dest='graph_users', action="store_true",default=False,                            
    #                         help="Expand graph users, don't summarise")
    group = parser.add_argument_group("creds","credential management").add_mutually_exclusive_group(required=False)
    group.add_argument('--credentials',dest='credentials', type=str, 
                       help='Credentials separated by column like bp1\user:password') 
    group.add_argument('--credfile',dest='credfile', type=str, 
                       help='Credential file which contains one line with rot13 + base64 encoded credentials like bp1\user:password') 
    global args    
    args=parser.parse_args()
    # This is for py2exe to work:
    if win32com.client.gencache.is_readonly == True:
    
        #allow gencache to create the cached wrapper objects
        win32com.client.gencache.is_readonly = False
    
        # under p2exe the call in gencache to __init__() does not happen
        # so we use Rebuild() to force the creation of the gen_py folder
        win32com.client.gencache.Rebuild(int(args.debug))
    
        # NB You must ensure that the python...\win32com.client.gen_py dir does not exist
        # to allow creation of the cache in %temp%
    if args.format_type=="graph-dot" and args.output is None:
        print >>sys.stderr, "ERROR: If you specify graph-dot, you must specify output!\n\n"
        parser.print_help()
        sys.exit(-1)
    
    # if args.format_type=="csv" and args.output:
    #     ext=args.output.split(".")[-1]
    #     if not ext.lower().endswith("csv"):
    #         logging.warning("You requested CSV output, but specified output file with a different extension: %s" % ext)
    level=logging.INFO
    if args.debug:        
        level=logging.DEBUG
        
    # if args.format_type=="csv" and args.direction=="both":
    #     print >>sys.stderr, "Direction 'both' is not supported for CSV"
    #     sys.exit(-1)
    
    logging.basicConfig(level=level,stream=sys.stderr)
    output=sys.stdout
    if args.output and args.format_type!="graph-dot":
        output=file(args.output,"wb")                   




    if args.credentials:
        (username,password)=args.credentials.split(":")
    elif args.credfile:
        (username,password)=file(args.credfile,'rb').read().decode('base64').decode('rot13').split(":")
    else:
        username=None
        password=None    
    global datasource
    if args.source.upper().startswith("LDAP"):
        datasource=ldapaccess.LDAPAccess(args.source,username,password)
        root="" # Cheating
    else:
        datasource=adsilib.UsefulADSISearches(args.source,username,password)
        root=datasource.getDefaultNamingContenxt()

    # fulltree=_recurse(tree,args.recursive)
    # connections=_buildConnections(datasource)  
    # if args.format_type=="pretty-print":
    #     pprint.PrettyPrinter(stream=output,indent=2,depth=40).pprint(connections)

    # elif args.format_type=="csv":
    #     raise NotImplementedError("Sorry, this feature is not implented yet!")
    #     # writer=csv.writer(output,quoting=csv.QUOTE_MINIMAL)        
    #     # _write_csv_nested(writer,[],fulltree)

    if args.format_type=="graph-dot" or args.format_type=="graph":
        # print "lalalal"
        graph=pydot.Dot(graph_type='digraph',fontname="Verdana",
                        rankdir='LR'#,label="Direction: %s" % args.direction,labelloc="top"
                        )

        _build_graph(graph,getSites(datasource),getSiteLinks(datasource))

        # Add legend
#         legend=pydot.Node("legend",shape="none",margin="0",label=\
#     """
# <<TABLE BORDER="0" CELLBORDER="1" CELLSPACING="0" CELLPADDING="4">    
#      <tr><td colspan="2">Legend</td></tr>
#      <tr><td colspan="2">Site replication map</td></tr>
# <TR>
#       <TD> Regular group </TD>
#       <TD BGCOLOR="#bfefff"> &nbsp;&nbsp;&nbsp;&nbsp;<br/></TD>
# </TR>
#      <TR>
#       <TD>Group contributing to multiple inheritance</TD>
#       <TD BGCOLOR="#ff69b4"> &nbsp;&nbsp;&nbsp;&nbsp;</TD>
#      </TR>
#      <TR>
#       <TD>Mutually inherited groups</TD>
#       <TD BGCOLOR="#ff6a6a"> &nbsp;&nbsp;&nbsp;&nbsp;</TD>
#      </TR>
#     </TABLE>>""")
#         graph.add_node(legend)
        #Monkeypatch pydot to convert unicode to utf8
        graph.mp_to_string=graph.to_string
        
        def _encodeme():
            try:
                return graph.mp_to_string().encode('utf8')
            except UnicodeDecodeError:
                return graph.mp_to_string()
        graph.to_string=_encodeme

        try:
            if args.format_type == "graph":

                (fd,tmpf)=tempfile.mkstemp()
                logging.debug("Tempfile is %s" % tmpf)
                os.close(fd)

                graph.write(tmpf)

                ext=args.output.split(".")[-1]
                if not os.path.exists(args.dotfile):
                    logging.CRITICAL("dot.exe does not exist, I got path %s" % args.dotfile)
                    sys.exit(-1)
                callargs=[args.dotfile,"-T%s" % ext.lower(),"-o",args.output,tmpf]

                if args.debug:
                    callargs.append("-v")
                logging.debug("Call args: %s" % str(callargs))
                subprocess.call(callargs,shell=True)
            else:
                graph.write(args.output)
        finally:
            pass
                #os.unlink(tmpf)


# def _make_graph_group(group):
#     #n=pydot.Node(_filter_node_badchars(group['name']),shape="rect",color="lightblue2", style="filled",
#     n=pydot.Node(group['name'],shape="rect",color="lightblue2", style="filled",                 
#                  #label="%s\n(%s)" % (group['name'],group['attrs']['distinguishedName'])
#                  label="%s" % (group['name'],)
#                  )
#     return n
def _make_graph_site(site):
    #n=pydot.Node(_filter_node_badchars(group['name']),shape="rect",color="lightblue2", style="filled",
    n=pydot.Node(site['_DN'],shape="rect",color="lightblue2", style="filled",                 
                 #label="%s\n(%s)" % (group['name'],group['attrs']['distinguishedName'])
                 label="%s\n%s" % (site['cn'],site['location'])
                 )
    return n

# global alreadyGraphNodes
# alreadyGraphedNodes={}
def _build_graph(graph,sites,links):
    # Scan costs 
    costs=[]
    for link in links:
        cost=int(link['cost'])
        if not cost in costs:
            costs.append(cost)
    costs.sort(reverse=True)
    graphedLinks=[]
    for site in sites:
        graph.add_node(_make_graph_site(site))
    # logging.DEBUG("Graphed %i sites" % len(sites) )
    for link in links:
        
        if "DefaultSiteLink" in link['cn']:
            logging.log(logging.DEBUG, "Skipping %s" % link['cn'])
            continue
        else:
            logging.log(logging.DEBUG, "Plotting %s" % link['cn'])
        for firstLink in link['siteList']:
            for secondLink in link['siteList']:
                if firstLink == secondLink:
                    continue
                if (firstLink,secondLink) in graphedLinks or (secondLink,firstLink) in graphedLinks:
                    continue
                cost=int(link['cost'])

                weight=costs.index(cost)+1
                logging.log(logging.DEBUG,"Cost was %i, weight is %i" % (cost,weight))
                edge=pydot.Edge(firstLink,secondLink,style="bold",weight=weight,label="%s/%s" % (link['cost'],link['replInterval']))
                graph.add_edge(edge)
                graphedLinks.append((firstLink,secondLink))

    






if __name__=="__main__":
    main()