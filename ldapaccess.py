import ldap
import ldap.sasl
import re 
import unittest
import logging
from ldap.controls import SimplePagedResultsControl
import pywintypes,win32security
logger=logging.getLogger("LDAPAccess")
LDAP_PAGE_SIZE=250
class LDAPAccess(object):
    def __init__(self,uri,username=None,password=None):
        self.uri=uri
        self.username=username
        self.password=password
        self.conn=None
        m=re.search("^(LDAP|LDAPS)://([^/]+)/(.*)$",uri)
        if not m:
            raise ValueError("Invalid LDAP URI '%s'" % self.uri)
        (proto,host,base)=m.groups()
        self.base=base
        self.host=host
        self.proto=proto
        
    def _getConection(self):
        
        if self.conn is None:
            conn=ldap.initialize("%s://%s" % (self.proto,self.host))
            conn.set_option(ldap.OPT_REFERRALS, 0)
            if(self.username):
                try:
                    conn.simple_bind_s(self.username,self.password)
                    logger.debug("Connected to %s as %s" % (self.uri,self.username))
                except ldap.STRONG_AUTH_REQUIRED:
                    auth_tok=ldap.sasl.digest_md5(self.username,self.password)
                    conn.sasl_interactive_bind_s("",auth_tok)
                    logger.debug("Connected to %s as %s via SASL-MD5" % (self.uri,self.username))
            else:
                logger.debug("Connected to %s anonymously" % self.uri)
            self.conn=conn
        
        return self.conn
    
    
    def search(self,searchFilter,attributes=None,scope=ldap.SCOPE_SUBTREE,base=None,sizelimit=0):
        if base is None:
            base=self.base
        l=self._getConection()
        lc = SimplePagedResultsControl(True,size=LDAP_PAGE_SIZE,cookie='')
        results=[]
        known_ldap_resp_ctrls = {
            SimplePagedResultsControl.controlType:SimplePagedResultsControl,
        }
        msgid=l.search_ext(base,scope,searchFilter,attributes,sizelimit=sizelimit,serverctrls=[lc])
        pages =0 
        while True:
            pages +=1
            logger.debug("Getting page %i" % pages)
            rtype, rdata, rmsgid, serverctrls = l.result3(msgid,resp_ctrl_classes=known_ldap_resp_ctrls)
            logger.debug("Retrieved %i records" % len(rdata))
            

            for res in  rdata:
                (dn,attribs)=res
                if type(attribs)==list:
                    continue
                attribs['_DN']=dn
                for (k,v) in attribs.items():
                    if len(v)==1:
                        attribs[k]=v[0]
                results.append(attribs)
            pctrls = [
                c
                for c in serverctrls
                if c.controlType == SimplePagedResultsControl.controlType
              ]
            if pctrls:
                #est, cookie = pctrls[0].controlValue
                if pctrls[0].cookie:
                    lc.cookie = pctrls[0].cookie
                    msgid=l.search_ext(base,scope,searchFilter,attributes,sizelimit=0,serverctrls=[lc])
                else:
                    break
            else:
                logger.warn("Warning:  Server ignores RFC 2696 control.")
                break
        return results
    
    
    def findGroups(self,groupName,attributes=None):
        searchFilter="(&(objectClass=group)(|(distinguishedName=%s)(name=%s)))" % (_escape(groupName),_escape(groupName))
        return self.search(searchFilter,attributes)
    
    def findUserOrGroup(self,objname,attributes=None):
        searchFilter="(&(|(objectClass=group)(objectClass=user))(|(distinguishedName=%s)(sAMAccountName=%s)))" % (_escape(objname),_escape(objname))
        return self.search(searchFilter,attributes)

    def findGroupMembers(self,groupName,objectClasses=None,attributes=None):
        # just same basic sanity check:
        assert re.search("\w+=\w+,",groupName),\
               "The group name must be in DN format. You supplied: %s" % groupName
        searchFilter="(&(objectClass=group)(distinguishedName=%s))" % (groupName)        
        res=self.search(searchFilter,["objectSid"])
        #print res
        rawsid=res[0]['objectSid']
        objectSid=win32security.ConvertSidToStringSid(pywintypes.SID(rawsid))
        primaryGroupID=objectSid.split("-")[-1]
        logging.log(logging.DEBUG,"objectSid: %s, RID: %s" % (objectSid,primaryGroupID))    
        if type(objectClasses)==list and len(objectClasses)>0:
            searchFilter="(&(|(primaryGroupID=%s)(memberof=%s))(|%s))" % \
            (primaryGroupID,groupName,"".join(["(objectClass=%s)" % x for x in objectClasses]))                                
        else:
            searchFilter="(|(primaryGroupID=%s)(memberof=%s))" % _escape(groupName)
        return self.search(searchFilter,attributes)

    def findTrustedDomains(self,attributes=None):
        searchFilter='(objectClass=trustedDomain)'
        res=self.search(searchFilter,attributes)#,base="CN=System")
        return res

class TestADSISearcher(unittest.TestCase):
    _uri="LDAP://UKSSQGC01/DC=uk,DC=kworld,DC=kpmg,DC=com"
    #_creds=('uk\ksmelkovs1',"YXJ5aG8xZ09sZmdlYnc=\n".decode('base64').decode('rot13'))
    #_creds=(username,password)=file(r'creds.txt','rb').read().decode('base64').decode('rot13').split(":")
    def setUp(self):
        self.searcher=LDAPAccess(self._uri,*self._creds)
        
    def test_base(self):
        self.assertEquals(self.searcher.base,"DC=uk,DC=kworld,DC=kpmg,DC=com")
        self.assertEquals(self.searcher.proto,"LDAP")
        self.assertEquals(self.searcher.host,"UKSSQGC01")
    
    def test_getEntryFromAD(self):
        result=self.searcher.search("(name=Domain Admins)")
        self.assertTrue(type(result)==list)
        self.assertGreaterEqual(len(result),1)
        checkDict={u'name':'Domain Admins',u'objectClass':['top','group']}
        self.assertDictContainsSubset(checkDict,result[0])
        
    def test_getSelectedAttributes(self):
        result=self.searcher.search("(name=Domain Admins)",["name","objectClass"])
        self.assertGreaterEqual(len(result),1)
        checkDict={u'name':'Domain Admins',u'objectClass':['top','group']}
        self.assertDictContainsSubset(checkDict,result[0])
        self.assertNotIn(u'objectSid',result[0])    
        
    #def test_paging(self):
        #global LDAP_PAGE_SIZE
        #oldpagesize=LDAP_PAGE_SIZE
        #try:
            
            #LDAP_PAGE_SIZE=1
            #ldap=LDAPAccess(self._uri,*self._creds)
            #res=ldap.search("(|(name=Domain Admins)(name=Administrators))")
            #self.assertEqual(len(res),2)
            
        #finally:
            #LDAP_PAGE_SIZE=oldpagesize

def _escape(attr):
    from ldap import filter as f
    return f.escape_filter_chars(attr,1)