# AD group expansion application
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




sys.setrecursionlimit(3000)
def _filter_node_badchars(n):
    assert type(n)==str or type(n)==unicode
    return re.sub("[^A-Za-z0-9_]","_",n)
    #return n.replace("&- ","_")

def _makeTreeEntry(node,rel):
    assert rel in "forward reverse none".split(" "),\
           "Unknown relationship type between nodes %s" % rel
    tree={'name':node['name'],'attrs':node,'users':[],'groups':[],'rel':rel}
    return tree

def _find_graphviz_dot():
    candidates=glob.glob(os.path.join(os.environ['PROGRAMFILES'],"GraphViz*","bin","dot.exe"))
    if len(candidates):
        dotfile=candidates[0]
        if os.path.exists(dotfile):
            return dotfile
    return None
def main():
    format_types="csv pretty-print graph-dot".split(" ")
    #lookup_directions="reverse forward both".split(" ")
    lookup_directions="reverse forward".split(" ")
    parser = argparse.ArgumentParser(description="Retrieve and expand AD groups.\nIf no credentials are supplied current credentials of logged on user will be used",
                                     epilog=\
"""
Supported output formats are:
 * csv - dump to a CSV file
 * pretty-print - dump data structures in indented fashion (for debugging)
 * graph-dot - output a .dot file which is agraph defintion as understood by GraphViz.
 By Konrads Klints <konrads.klints@kpmg.co.uk>.
                                 """.strip())
    parser.add_argument('groupName',metavar='OBJECT',type=str,nargs=1,
                        help='Object (group or user) to expand. You can specify both DN and NT names')    
    parser.add_argument('--debug',dest='debug', action='store_true', default=False,
                        help='Enable debugging output')
    parser.add_argument('-o','--output',dest='output', type=str,metavar="FILE",
                        help='Output file, defaults to stdout') 
    parser.add_argument('--no-recursive',dest='recursive', action="store_false",default=True,
                        help='Do not expand child groups')   
    
    parser.add_argument('-f','--format',dest='format_type', type=lambda x: x.lower(),choices=format_types,
                        default=format_types[0],
                        help="Output format")
    parser.add_argument("-s","--data-source",
                        help="Specify data source. Defaults to global catalogue GC:",
                        type=str,default="GC:",dest='source')
    dumpwhat=parser.add_mutually_exclusive_group()  
    dumpwhat.add_argument("-d","--direction",
            help="""Lookup direction. Forward looks up the members of the group (only for groups), 
            reverse - groups that the specified object is a member of. 
            The default for groups is forward, while for users - reverse""",
            type=lambda x: x.lower(),choices=lookup_directions,
            dest='direction')
    dumpwhat.add_argument("-A","--dump-all",
                          help="Dump every group",
                          dest='dump_all',default=False,action="store_true")
    ugg=parser.add_argument_group("User expansion",
                                  """
Options related to user expansion. By default, if input object is group, then users will be graphed, 
if user, then not"""
                                  ).add_mutually_exclusive_group()    
    ugg.add_argument('--fetch-users',dest='always_fetch_users', action="store_true",    default=None ,                         
                     help="Always force users")
    ugg.add_argument('--no-fetch-users',dest='never_fetch_users', action="store_true", default=None,                           
                     help="Never fetch users")    
    csvgroup=parser.add_argument_group('csv','Arguments relating to CSV output')
    csvgroup.add_argument('--no-dn',dest='csv_output_dn', action="store_false",default=True,                            
                          help="Do not print DNs in CSV")
    graphgroup=parser.add_argument_group("graphing","Options related to graph generation")
    dotfile=_find_graphviz_dot()
    if dotfile is None:
        dotfile="NOT FOUND"
    graphgroup.add_argument('--graphviz-dot',dest='dotfile', type=str,default=dotfile,                            
                            help="Path to GraphViz dot.exe, best guess: %s" % dotfile)
    graphgroup.add_argument('--graph-users',dest='graph_users', action="store_true",default=False,                            
                            help="Expand graph users, don't summarise")
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
    
    if args.format_type=="csv" and args.output:
        ext=args.output.split(".")[-1]
        if not ext.lower().endswith("csv"):
            logging.warning("You requested CSV output, but specified output file with a different extension: %s" % ext)
    level=logging.INFO
    if args.debug:        
        level=logging.DEBUG
        
    if args.format_type=="csv" and args.direction=="both":
        print >>sys.stderr, "Direction 'both' is not supported for CSV"
        sys.exit(-1)
    
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
    if args.source.startswith("LDAP"):
        datasource=ldapaccess.LDAPAccess(args.source,username,password)
        root="" # Cheating
    else:
        datasource=adsilib.UsefulADSISearches(args.source,username,password)
        root=datasource.getDefaultNamingContenxt()
    global defaultAttrs
    defaultAttrs="distinguishedName,name,member,memberOf,objectClass".split(",")
    
    if args.dump_all:
        if not args.source.startswith("LDAP"):            
            print >>sys.stderr, "Dump all with non LDAP data source is not supported!"
            sys.exit(-1)
        args.fetch_users=None
        
        global dump_all_results    
        dump_all_results=datasource.search("objectClass=group",["name","distinguishedName","member","memberOf"])
        
        fakeRootNode=_makeTreeEntry({'name':'__fakeRoot','distinguishedName':'__fakeRootDN'},"none")
        args.direction="forward"
        fakeRootNode['attrs']['member']=dump_all_results
        logging.debug("Dump all returned %i nodes" % len(dump_all_results))
        global memberOfHash        
        memberOfHash={}
        
        for e in dump_all_results:
            if e.has_key("memberOf"):
                if type(e['memberOf'])==str or type(e['memberOf'])==unicode:
                    mofs=[e['memberOf']]
                else:
                    mofs=e['memberOf']
                for m in mofs:
                   # dn=m['distinguishedName']
                    if not memberOfHash.has_key(m):
                        memberOfHash[m]=[]
                    memberOfHash[m].append(e)
        
        fulltree=_recurse(fakeRootNode,True)
        
    else:
        groupEntry=datasource.findUserOrGroup(args.groupName[0],defaultAttrs)
    
        if len(groupEntry)==0:
            logging.log(logging.INFO,"No groups named `%s` have been found" % args.groupName[0])
            return -1
        groupEntry=filter(lambda x: x['distinguishedName'].endswith(root),groupEntry)
    
        if len(groupEntry)>1:
            logging.log(logging.ERROR,"%i groups found named `%s`: %s" % \
                        (len(groupEntry),args.groupName[0],", ".join([g['distinguishedName'] for g in groupEntry]))
                        )
            return -1
        singleGroupEntry=groupEntry[0]
    
        #
        args.fetch_users=None
        if "user" in singleGroupEntry['objectClass']:
            if args.always_fetch_users is None and  args.never_fetch_users is None:
                args.fetch_users=False
                logging.debug("I got a user to expand and by default I am not fetching users")        
        elif "group" in singleGroupEntry['objectClass']:
            if args.always_fetch_users is None and  args.never_fetch_users is None:
                args.fetch_users=True
                logging.debug("I got a group to expand and by default I am fetching users")
        else:
            logging.error("Some kind of wrong entry?!")
            sys.exit(-1)    
    
        # if args.fetch_users is None:
        if args.always_fetch_users is not None:
            args.fetch_users=True

        if args.never_fetch_users is not None:
            args.fetch_users=False            
        assert args.fetch_users is not None    
        logging.debug("Ultimately, fetch_users is %s" % str(args.fetch_users))
        
        if "user" in singleGroupEntry['objectClass']:
            if args.direction is None:
                args.direction="reverse"
            elif args.direction=="forward":
                print >>sys.stderr, "ERROR: Forward lookup is an unsupported direction mode for user object!"
                sys.exit(-1)
        
        if "group" in singleGroupEntry['objectClass'] and args.direction is None:
            args.direction="forward"
        
        logging.debug("Search direction is %s" % args.direction)
            
        groupDN=singleGroupEntry['distinguishedName']
        logging.debug("Got group with DN: %s" % groupDN)
        tree=_makeTreeEntry(singleGroupEntry,"none")
    
    
    
    
        fulltree=_recurse(tree,args.recursive)
    if args.format_type=="pretty-print":
        pprint.PrettyPrinter(stream=output,indent=2,depth=40).pprint(fulltree)


    elif args.format_type=="csv":
        writer=csv.writer(output,quoting=csv.QUOTE_MINIMAL)        
        _write_csv_nested(writer,[],fulltree)

    elif args.format_type=="graph-dot":
        graph=pydot.Dot(graph_type='digraph',fontname="Verdana",
                        rankdir='LR'#,label="Direction: %s" % args.direction,labelloc="top"
                        )

        _build_graph(graph,fulltree)
        # Add legend
        legend=pydot.Node("legend",shape="none",margin="0",label=\
("""
<<TABLE BORDER="0" CELLBORDER="1" CELLSPACING="0" CELLPADDING="4">    
     <tr><td colspan="2">Legend</td></tr>
     <tr><td colspan="2">Direction: %s</td></tr>
<TR>
      <TD> Regular group </TD>
      <TD BGCOLOR="#bfefff"> &nbsp;&nbsp;&nbsp;&nbsp;<br/></TD>
</TR>
     <TR>
      <TD>Group contributing to multiple inheritance</TD>
      <TD BGCOLOR="#ff69b4"> &nbsp;&nbsp;&nbsp;&nbsp;</TD>
     </TR>
     <TR>
      <TD>Mutually inherited groups</TD>
      <TD BGCOLOR="#ff6a6a"> &nbsp;&nbsp;&nbsp;&nbsp;</TD>
     </TR>
    </TABLE>>""" % args.direction.upper()).strip())
        graph.add_node(legend)
        #Monkeypatch pydot to convert unicode to utf8
        graph.mp_to_string=graph.to_string
        
        def _encodeme():
            try:
                return graph.mp_to_string().encode('utf8')
            except UnicodeDecodeError:
                return graph.mp_to_string()
        graph.to_string=_encodeme



        
        try:
            (fd,tmpf)=tempfile.mkstemp()
            logging.debug("Tempfile is %s" % tmpf)
            os.close(fd)

            graph.write(tmpf)

            ext=args.output.split(".")[-1]
            if not os.path.exists(args.dotfile):
                logging.CRITICAL("dot.exe does not exist, I got path %s" % args.dotfile)
                sys.exit(-1)
            callargs=[args.dotfile,"-T%s" % ext.lower(),"-o",args.output.replace("&","-"),tmpf]

            if args.debug:
                callargs.append("-v")
            logging.debug("Call args: %s" % str(callargs))
            subprocess.call(callargs,shell=True)
        finally:
            pass
                #os.unlink(tmpf)


def _make_graph_group(group):
    #n=pydot.Node(_filter_node_badchars(group['name']),shape="rect",color="lightblue2", style="filled",
    n=pydot.Node(group['name'],shape="rect",color="lightblue2", style="filled",                 
                 #label="%s\n(%s)" % (group['name'],group['attrs']['distinguishedName'])
                 label="%s" % (group['name'],)
                 )
    return n
def _make_graph_user(user):
    n=pydot.Node(_filter_node_badchars(user['name']),shape="oval",
                 label="%s" % (user['name'],)
                 )
    return n    
global alreadyGraphNodes
alreadyGraphedNodes={}
def _build_graph(graph,tree,parent=None,level=0):
    name=tree['name']
    if name in alreadyGraphedNodes: # Detect a closed walk
        myself=alreadyGraphedNodes[name]
        color="hotpink"
        if filter(lambda x: x['name'] == parent.get_name().strip('"'),tree['groups']):
            color="indianred1"
            logging.debug("Possible loop detected between %s and %s" % (name,parent.get_name()))
        if not myself.get_color()=="indianred1":
            myself.set_color(color)
        if not parent.get_color()=="indianred1":
            parent.set_color(color)
        edge=pydot.Edge(parent,myself,style="bold")
        edge.set_color(color)
        graph.add_edge(edge)
        return myself

    assert isinstance(graph,pydot.Dot)
    
    myself=_make_graph_group(tree)
    if name!="__fakeRoot":
        graph.add_node(myself)
    if (parent is not None) and (parent.get_label().strip('"')!="__fakeRoot"):
        graph.add_edge(pydot.Edge(parent,myself))   
        #if args.direction=="reverse":
            #graph.add_edge(pydot.Edge(parent,myself))        
        #elif args.direction=="forward":
            #graph.add_edge(pydot.Edge(myself,parent))
            #myself.set_color("gainsboro")
    if len(tree['users']):

        #pydot.Subgraph
        cluster=pydot.Cluster(graph_name="users_%s" % _filter_node_badchars(tree['name']), 
                              label="Users of %s" % tree['name'],
                              )
        graph.add_subgraph(cluster)
        i=0;
        old_invis_node=None
        if args.graph_users:
            for u in tree['users']:

                userNode=_make_graph_user(u)
                cluster.add_node(userNode)            
                edge=pydot.Edge(myself,userNode)
                graph.add_edge(edge)
                #if i%5==0:                
                    #invis_node=pydot.Node("invis_%s_%i" % (tree['name'].replace(" ","_"),i),style="invis")
                    #cluster.add_node(invis_node)
                    #graph.add_edge(pydot.Edge(myself,invis_node,style="invis"))
                    #if old_invis_node:
                        #graph.add_edge(pydot.Edge(old_invis_node,invis_node,style="invis"))
                    #old_invis_node=invis_node
                #graph.add_edge(pydot.Edge(invis_node,userNode,style="invis"))
                #i+=1
        else:
            a=[x['name'] for x in tree['users']]
            empty_for_none=lambda x: [x,""][int((x is None))]
            users=",\n".join(map(lambda x,y: ", ".join([empty_for_none(x),empty_for_none(y)]),a[::2],a[1::2]))
            userNode=n=pydot.Node(users,shape="rect",color="invis",
                                  label="%s" % (users,)
                                  )
            cluster.add_node(userNode)            
            edge=pydot.Edge(myself,userNode)
            graph.add_edge(edge)

    alreadyGraphedNodes[name]=myself    
    for g in tree['groups']:
        _build_graph(graph,g,myself)

def _make_csv_entry(x):
    try:
        dn=x['attrs']['distinguishedName']
    except KeyError:
        dn=x['distinguishedName']
    if args.csv_output_dn:
        s="%s (%s)" % (x['name'],dn)
    else:
        s=x['name']
    return s.encode('utf8')

def _write_csv_nested(writer,path,node):
    name=node['name']
    if name in alreadyGraphedNodes:
        return
    prefixer = lambda x: map(lambda x: [x,' '+x][bool(x.startswith("-"))],x)
    assert type(args.csv_output_dn)==bool

    newpath=[]
    if name!="__fakeRoot":
        newpath.extend(path)
        newpath.append(_make_csv_entry(node))

    for u in node['users']:
        row=[]
        row.extend(newpath)        
        row.append(_make_csv_entry(u))        
        writer.writerow(prefixer(row))
    alreadyGraphedNodes[name]=1
    for g in node['groups']:
        row=[]
        row.extend(newpath)
        row.append(_make_csv_entry(g))
        writer.writerow(prefixer(row))
        _write_csv_nested(writer,newpath,g)
    #for entity in resultForOutput:
            #dn=entity['distinguishedName']
            #for (key,values) in entity.items():
                #if key in ["mSMQDigests"]:
                    #continue
                #if not (type(values)==list or type(values)==tuple):
                    #values=[values]
                #for v in list(values):
                    #try:
                        #writer.writerow([dn,key,v])                                   
                    #except ValueError,e:
                        #if e.message=="can't format dates this early":
                            #logging.debug("Ignoring early time: %s=%r in %s" % (key,v,dn))
                        #else:
                            #raise e
global expandedGroups
expandedGroups={}   

def _recurse(node,recurse=True,resolvedCount=0):
    logging.debug("Resolving %s (resolved %i so far)" % (node['name'],resolvedCount))
    if node['name'] in expandedGroups:
        logging.debug("Found group %s in already expanded list" % node['name'])
        return expandedGroups[node['name']]
    
    dn=node['attrs']['distinguishedName']
    if args.fetch_users:
        try:
            node['users'].extend(datasource.findGroupMembers(dn,["user"],defaultAttrs)) 
        except adsilib.ADSISearcherException,e:
            if e.code==adsilib.ADSISearcherException.SIZE_EXCEEDED:
                logging.warning("Too many users in group %s!" % dn)
            else:
                raise e

    expandedGroups[node['name']]=node
    if args.dump_all:
        members=None
        if dn=="__fakeRootDN":
            members=dump_all_results
        elif 'member' in node['attrs']:
                #members=node['attrs']['member']
                #members=filter(lambda x: x.has_key('memberOf') and dn in x['memberOf'],dump_all_results)
                if memberOfHash.has_key(dn):
                    members=memberOfHash[dn]
                else:
                    members=None
            
                #datasource.findGroupMembers(dn,["group"])
        if members is not None:
            if not (type(members)==list or type(members)==tuple):
                members=[members]
            for g in members:                                
                gnode=_makeTreeEntry(g,"forward")
                if recurse:
                    gnode=_recurse(gnode,resolvedCount+1)
                node['groups'].append(gnode)
    else:
        
        if args.direction=="both" or args.direction=="reverse":
            if 'memberOf' in node['attrs']:
                members=node['attrs']['memberOf']
                if members is not None:
                    if not (type(members)==list or type(members)==tuple):
                        members=[members]
                    for g in members:
                        try:
                            grp=datasource.findGroups(g)[0]
                        except IndexError:
                            continue
                        gnode=_makeTreeEntry(grp,"reverse")
                        if recurse:
                            gnode=_recurse(gnode,resolvedCount+1)
                        node['groups'].append(gnode)
                        
        if args.direction=="both" or args.direction=="forward":
            if 'member' in node['attrs']:
                #members=node['attrs']['member']
                members=datasource.findGroupMembers(dn,["group"])
                if members is not None:
                    if not (type(members)==list or type(members)==tuple):
                        members=[members]
                    for g in members:                                
                        gnode=_makeTreeEntry(g,"forward")
                        if recurse:
                            gnode=_recurse(gnode,resolvedCount+1)
                        node['groups'].append(gnode)
                    


    return node






if __name__=="__main__":
    main()