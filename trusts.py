# Plot trusts in an AD forest
# Author: Konrads Klints <konrads.klints@kpmg.co.uk>
import sys
import logging
import adsilib
import ldap
import csv
import argparse
import win32com
import win32com.client as win32
import collections
import pprint
import pydot
import os,glob,tempfile,subprocess,re
import ldapaccess
import dns.resolver


class Enough(Exception):
    pass
sys.setrecursionlimit(3000)
def get_dcs_for_domain(domain):
    fqdn="_ldap._tcp.dc._msdcs.%s" % domain
    logging.debug("Trying to query " + fqdn)
    answers = dns.resolver.query(fqdn,'SRV')
    return map(lambda ans: str(ans.target),answers)
def dnstoldap(dns):
    segments=dns.split(".")
    return ",".join(map(lambda x: "DC=" +x,segments))

def getLdapSourceFromDomain(domain,username,password):
        
        servers=get_dcs_for_domain(domain)
        if len(servers)==0:
            print >>sys.stderr, "Couldn't get a DC for domain " + domain
            sys.exit(-1)
        servers=map(lambda x: x.rstrip('.'),servers)
        for server in servers:
            try:
                ldapsource="LDAP://%s/%s" % (server,dnstoldap(domain))
                # ldapsource="LDAP://%s/%s" % (server,"Default naming context")
                logging.debug("The constructed LDAP source is: '%s'" % ldapsource)
                acc= ldapaccess.LDAPAccess(ldapsource,username,password)
                acc._getConection()
                break
            except (ldap.SERVER_DOWN,ldap.AUTH_UNKNOWN, ldap.STRONG_AUTH_REQUIRED) as e:
                logging.debug( "Can't connect to %s: %s" % (server,str(e)))
                if e.message['desc']=="Unknown authentication method":
                    logging.debug( "Can't connect to %s: %s, skipping the whole lot" % (server,str(e)))
                    acc=None
                    break                    
                acc=None
                
            except Exception as e:
                logging.debug( "Bugger on %s :( %s" % (server,str(e)))     
                acc=None   
        return acc
         

def _filter_node_badchars(n):
    assert type(n)==str or type(n)==unicode
    return re.sub("[^A-Za-z0-9_]","_",n)
    #return n.replace("&- ","_")

TRUSTDIRECTION={
    0: "DISABLED",
    2: "OUTBOUND",
    3: "BOTH"
}

def _makeTreeEntry(domain):    
    tree={'domain':domain, 'trusts':{},'errflag':None}
    # tree={'name':node['name'],'attrs':node,'trusts':[],'groups':[],'rel':rel}
    return tree



def _formatTrusts(trusts):

    treeTrusts={}
    for t in trusts:
        e={}
        e['domain']=t['name']
        e['trustDirection']=t['trustDirection']
        e['trustDirectionText']=TRUSTDIRECTION[int(t['trustDirection'])]
        e['trustAttributes']=t['trustAttributes']
        e['node']=None
        treeTrusts[t['name']]=e
    return treeTrusts

def _find_graphviz_dot():
    candidates=glob.glob(os.path.join(os.environ['PROGRAMFILES'],"GraphViz*","bin","dot.exe"))
    if len(candidates):
        dotfile=candidates[0]
        if os.path.exists(dotfile):
            return dotfile
    return None
def main():
    #format_types="csv pretty-print graph-dot".split(" ")
    format_types="pretty-print graph-dot".split(" ")
    #lookup_directions="reverse forward both".split(" ")
    lookup_directions="reverse forward".split(" ")
    parser = argparse.ArgumentParser(description="Plot domain trusts.\nIf no credentials are supplied current credentials of logged on user will be used",
                                     epilog=\
"""
Supported output formats are:
 * csv - dump to a CSV file (disabled)
 * pretty-print - dump data structures in indented fashion (for debugging)
 * graph-dot - output a .dot file which is agraph defintion as understood by GraphViz.
 By Konrads Klints <konrads.klints@kpmg.co.uk>.
                                 """.strip())
    parser.add_argument('startDomain',metavar='DOMAIN',type=str,nargs=1,
                        help='Domain (DNS style) with which to start')    
    parser.add_argument('--debug',dest='debug', action='store_true', default=False,
                        help='Enable debugging output')
    parser.add_argument('-o','--output',dest='output', type=str,metavar="FILE",
                        help='Output file, defaults to stdout') 
    parser.add_argument('--no-recursive',dest='recursive', action="store_false",default=True,
                        help='Do not plot trusts of other domains')   
    parser.add_argument("-s","--data-source",
                        help="Specify data source. Defaults to DNS resolution using SRV records",
                        type=str,default="DNS",dest='source')
    parser.add_argument("--skip-domains",
                            help="Skip some domains", 
                            dest="skip",
                            type=str,default="")
    
    parser.add_argument('-f','--format',dest='format_type', type=lambda x: x.lower(),choices=format_types,
                        default=format_types[0],
                        help="Output format")
    
    graphgroup=parser.add_argument_group("graphing","Options related to graph generation")
    dotfile=_find_graphviz_dot()
    if dotfile is None:
        dotfile="NOT FOUND"
    graphgroup.add_argument('--graphviz-dot',dest='dotfile', type=str,default=dotfile,                            
                            help="Path to GraphViz dot.exe, best guess: %s" % dotfile)
    # graphgroup.add_argument('--graph-users',dest='graph_users', action="store_true",default=False,                            
                            # help="Expand graph users, don't summarise")
    group = parser.add_argument_group("creds","credential management").add_mutually_exclusive_group(required=False)
    group.add_argument('--credentials',dest='credentials', type=str, 
                       help='Credentials separated by column like uk\user:password') 
    group.add_argument('--credfile',dest='credfile', type=str, 
                       help='Credential file which contains one line with rot13 + base64 encoded credentials like bp1\user:password') 
    global args    
    args=parser.parse_args()
    args.skip=args.skip.split(",")
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
        
    
    logging.basicConfig(level=level,stream=sys.stderr)
    output=sys.stdout
    if args.output and args.format_type!="graph-dot":
        output=file(args.output,"wb")                   


    # fulltree=None
    global username
    global password
    if args.credentials:
        (username,password)=args.credentials.split(":")
    elif args.credfile:
        (username,password)=file(args.credfile,'rb').read().decode('base64').decode('rot13').split(":")
    else:
        username=None
        password=None    
        
    trusts={}
    if args.source.startswith("LDAP"):
        datasource=ldapaccess.LDAPAccess(args.source,username,password)
        root="" # Cheating
        if args.recursive:
            print >>sys.stderr("Recursive mode with single LDAP source doesn't make sense")
            sys.exit(-1)
    elif args.source=="DNS":
        pass
        # datasource=getLdapSourceFromDomain(args.startDomain[0],username,password)
    
    elif args.source=="pretty-print":
        try:
            
            logging.debug("Loading pretty-print source: %s" % args.startDomain[0])                
            trusts=eval(file(args.startDomain[0],"r").read())
        
        except OSError,e:
            print >>sys.stderr, "Could not load pretty-print source" + args.startDomain[0]
            sys.exit(-1)
    else:
        print >>sys.stderr, "Unknown data source " + args.source
        sys.exit(-1)


    #begin constructing tree
    global defaultAttrs
    defaultAttrs="cn,trustDirection,flatName,name,trustType,trustAttributes,trustDirection,objectClass".split(",")           

    
    if len(trusts)==0: # pretty-print already fills this data in
        pendingDomains=[args.startDomain[0]]
        alreadyTraversedDomains=[]
    
        while len(pendingDomains)>0:
            domain=pendingDomains.pop()
            if domain in alreadyTraversedDomains:
                logging.debug("Skipping over %s as we've already traversed it" % domain)
                continue
            
                
            alreadyTraversedDomains.append(domain)
    
            trusts[domain]={'domain' : domain, 'errflag' : None, 'trusts': {}} # init struct
            if domain in args.skip:
                logging.debug("Skipping over %s as per skip-list" % domain)
                trusts[domain]['errflag']='SKIPPED'
                continue
    
            datasource=getLdapSourceFromDomain(domain,username,password)        
            if datasource is None:  # if we couldn't get the ldap connection for some reason
                trusts[domain]['errflag']="CANT_GET_TRUSTS"
                continue
            domainTrusts=datasource.findTrustedDomains(defaultAttrs)
            trusts[domain]['trusts']= _formatTrusts(domainTrusts)
    
            for trust in trusts[domain]['trusts'].values():
                if trust['domain'] not in pendingDomains and trust['domain'] not in alreadyTraversedDomains:
                    pendingDomains.append(trust['domain'])
            
        # end constructing tree

    if args.format_type=="pretty-print":
        pprint.PrettyPrinter(stream=output,indent=2,depth=40).pprint(trusts)


    elif args.format_type=="csv":
        raise NotImplementedError("CSV not yet implemented")
        # writer=csv.writer(output,quoting=csv.QUOTE_MINIMAL)        
        # _write_csv_nested(writer,[],fulltree)

    elif args.format_type=="graph-dot":
        graph=pydot.Dot(graph_type='digraph',fontname="Verdana",
                        rankdir='LR'#,label="Direction: %s" % args.direction,labelloc="top"
                        )

        _build_graph(graph,trusts)
        # Add legend
        legend=pydot.Node("legend",shape="none",margin="0",label="hello world")
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


# def _make_graph_group(group):
#     #n=pydot.Node(_filter_node_badchars(group['name']),shape="rect",color="lightblue2", style="filled",
#     n=pydot.Node(group['name'],shape="rect",color="lightblue2", style="filled",                 
#                  #label="%s\n(%s)" % (group['name'],group['attrs']['distinguishedName'])
#                  label="%s" % (group['name'],)
#                  )
#     return n
# def _make_graph_user(user):
#     n=pydot.Node(_filter_node_badchars(user['name']),shape="oval",
#                  label="%s" % (user['name'],)
#                  )
#     return n    

def _make_graph_domain(domain):
    
    color="lightblue2"
    style="filled"
    if domain['errflag']=='SKIPPED':
        color="lightgrey"
        style="dashed"
    if domain['errflag']=='CANT_GET_TRUSTS':
        color="indianred"
    n=pydot.Node(_filter_node_badchars(domain['domain']),
        shape="rect",
        style=style,
        label="%s" % (domain['domain']),
        color=color)

    return n    

def _make_trust_edge(domainName,trust):
    label="Direction: %s, attribtues: %s" % (trust['trustDirectionText'],trust['trustAttributes'])
    if trust['trustDirectionText']=='BOTH':
        direction="both"
    else:
        direction="forward"
    e=pydot.Edge(_filter_node_badchars(domainName),_filter_node_badchars(trust['domain']),label=label,dir=direction)
    return e

    


def _build_graph(graph,tree):
    
    alreadyGraphedTrusts=[]
    for domainName, domain in tree.items():
        node=_make_graph_domain(domain)
        graph.add_node(node)

    for domainName, domain in tree.items():
        for t in domain['trusts'].values():
            domainName=domain['domain']
            x=[domainName,t['domain']]
            sorted(x)
            #trustTuple=
            if x not in alreadyGraphedTrusts:
                graph.add_edge(_make_trust_edge(domainName,t))
                alreadyGraphedTrusts.append(x)

    # name=tree['domain']
    # if name in alreadyGraphedNodes: # Detect a closed walk        
    #     logging.debug("We've already graphed trusts of %s" % name)
    #     return tree
    # assert isinstance(graph,pydot.Dot)
    
    # myself=_make_graph_group(tree)
    # if name!="__fakeRoot":
    #     graph.add_node(myself)
    # if (parent is not None) and (parent.get_label().strip('"')!="__fakeRoot"):
    #     graph.add_edge(pydot.Edge(parent,myself))   
    #     #if args.direction=="reverse":
    #         #graph.add_edge(pydot.Edge(parent,myself))        
    #     #elif args.direction=="forward":
    #         #graph.add_edge(pydot.Edge(myself,parent))
    #         #myself.set_color("gainsboro")
    # if len(tree['users']):

    #     #pydot.Subgraph
    #     cluster=pydot.Cluster(graph_name="users_%s" % _filter_node_badchars(tree['name']), 
    #                           label="Users of %s" % tree['name'],
    #                           )
    #     graph.add_subgraph(cluster)
    #     i=0;
    #     old_invis_node=None
    #     if args.graph_users:
    #         for u in tree['users']:

    #             userNode=_make_graph_user(u)
    #             cluster.add_node(userNode)            
    #             edge=pydot.Edge(myself,userNode)
    #             graph.add_edge(edge)
    #             #if i%5==0:                
    #                 #invis_node=pydot.Node("invis_%s_%i" % (tree['name'].replace(" ","_"),i),style="invis")
    #                 #cluster.add_node(invis_node)
    #                 #graph.add_edge(pydot.Edge(myself,invis_node,style="invis"))
    #                 #if old_invis_node:
    #                     #graph.add_edge(pydot.Edge(old_invis_node,invis_node,style="invis"))
    #                 #old_invis_node=invis_node
    #             #graph.add_edge(pydot.Edge(invis_node,userNode,style="invis"))
    #             #i+=1
    #     else:
    #         a=[x['name'] for x in tree['users']]
    #         empty_for_none=lambda x: [x,""][int((x is None))]
    #         users=",\n".join(map(lambda x,y: ", ".join([empty_for_none(x),empty_for_none(y)]),a[::2],a[1::2]))
    #         userNode=n=pydot.Node(users,shape="rect",color="invis",
    #                               label="%s" % (users,)
    #                               )
    #         cluster.add_node(userNode)            
    #         edge=pydot.Edge(myself,userNode)
    #         graph.add_edge(edge)

    # alreadyGraphedNodes[name]=myself    
    # for g in tree['groups']:
    #     _build_graph(graph,g,myself)

# def _make_csv_entry(x):
#     try:
#         dn=x['attrs']['distinguishedName']
#     except KeyError:
#         dn=x['distinguishedName']
#     if args.csv_output_dn:
#         s="%s (%s)" % (x['name'],dn)
#     else:
#         s=x['name']
#     return s.encode('utf8')

# def _write_csv_nested(writer,path,node):
#     name=node['name']
#     if name in alreadyGraphedNodes:
#         return
#     prefixer = lambda x: map(lambda x: [x,' '+x][bool(x.startswith("-"))],x)
#     assert type(args.csv_output_dn)==bool

#     newpath=[]
#     if name!="__fakeRoot":
#         newpath.extend(path)
#         newpath.append(_make_csv_entry(node))

#     for u in node['users']:
#         row=[]
#         row.extend(newpath)        
#         row.append(_make_csv_entry(u))        
#         writer.writerow(prefixer(row))
#     alreadyGraphedNodes[name]=1
#     for g in node['groups']:
#         row=[]
#         row.extend(newpath)
#         row.append(_make_csv_entry(g))
#         writer.writerow(prefixer(row))
#         _write_csv_nested(writer,newpath,g)
  
# global expandedGroups
# expandedGroups={}   

# def _recurse(node,resolvedCount=0):           
#     global username
#     global password
#     global args
#     print(node)
#     logging.debug("Resolving %s (resolved %i so far)" % (node['domain'],resolvedCount))
#     if node['domain'] in expandedGroups:
#         logging.debug("Found group %s in already expanded list" % node['domain'])
#         return expandedGroups[node['domain']]    
#     expandedGroups[node['domain']]=node
#     if node['domain'] in args.skip:
#         return node
#     datasource=getLdapSourceFromDomain(node['domain'],username,password)
#     if datasource is None:
#         node['errflag']="CANT_GET_TRUSTS"
#         return node
#     trusts=datasource.findTrustedDomains(defaultAttrs)
#     node['trusts']=formatTrusts(trusts)
   
#     for n,t in node['trusts'].items():
#         resolvedCount+=1
#         t['node']=_recurse(_makeTreeEntry(n),resolvedCount)       
           
#         # newnode=_makeTreeEntry(t['domain'],trusts) # one of the trust, i will recurse this next
#         # t['node'],t['domain']=newnode                                
#         # _recurse(t,resolvedCount+1)
#     #for t in newnode['tusts']:                         
#     return node

if __name__=="__main__":
    main()