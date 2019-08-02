# Based on search.py of win32comext demo, Python PSF License
# Author: Konrads Klints <konrads.klints@kpmg.co.uk>
from win32com import adsi
from win32com.adsi.adsicon import *
from win32com.adsi import adsicon
import pythoncom, pywintypes, win32security
import logging
import unittest
import re

class ADSISearcherException(Exception):
    SIZE_EXCEEDED=1    
    def __init__(self,*args,**kwargs):        
        self.code=kwargs.get('code',None)
        super(ADSISearcherException,self).__init__(*args)
class ADSISearcher(object):
    def __init__(self,source="GC:",username=None,password=None):
        self.username=username
        self.password=password
        self.source=source
        self.gc=None
        self.defaultADSIObject=None
        self.converters = {
        'objectGUID' : self._guid_from_buffer,
        'objectSid' : self._sid_from_buffer,
        'instanceType' :self.getADsTypeName,
    }
        
    ADsTypeNameMap = {}
    
    @classmethod
    def getADsTypeName(klass,type_val):
        # convert integer type to the 'typename' as known in the headerfiles.
        if not klass.ADsTypeNameMap:
            for n, v in adsicon.__dict__.iteritems():
                if n.startswith("ADSTYPE_"):
                    klass.ADsTypeNameMap[v] = n
        return klass.ADsTypeNameMap.get(type_val, hex(type_val))
    
    def _guid_from_buffer(self,b):
        return pywintypes.IID(b, True)
    
    def _sid_from_buffer(self,b):
        return str(pywintypes.SID(b))
    
    _null_converter = lambda self,x: x
    
   
    
    def log(level, msg, *args):
        logging.debug(msg % args)
        
    
    def getGC(self,force=False):
        
        if force is False and self.gc:
            return self.gc
        path=self.source
        logging.debug("Search path is: %s" % path)
        #cont = adsi.ADsOpenObject("LDAP://%s" % forestRoot, self.username, self.password, 0, adsi.IID_IADsContainer)
        cont = adsi.ADsOpenObject(path, self.username, self.password, 0, adsi.IID_IADsContainer)
        enum = adsi.ADsBuildEnumerator(cont)
        # Only 1 child of the global catalog.
        for e in enum:
            gc = e.QueryInterface(adsi.IID_IDirectorySearch)
            self.gc=gc
            return gc
        raise ADSISearcherException("Unable to get Global Catalogue")
    
    def getDefaultNamingContenxt(self):
        rootdse = adsi.ADsGetObject("LDAP://rootDSE")
        return rootdse.Get("defaultNamingContext")
    
    #def getDefaultADSIObject(self,force=False):
        #if force is False and self.defaultADSIObject:
            #return self.defaultADSIObject
        #rootdse = adsi.ADsGetObject("LDAP://rootDSE")
        #path="LDAP://" + rootdse.Get("defaultNamingContext")
        #adsi.ADsGetObject(path)
        #enum = adsi.ADsBuildEnumerator(adsi.ADsGetObject(path))
        ## Only 1 child of the global catalog.
        #for e in enum:
            #obj = e.QueryInterface(adsi.IID_IDirectorySearch)
            #self.defaultADSIObject=obj
            #return obj
        #raise ADSISearcherException("Unable to get default LDAP root object")
        
    def _convert_attribute(self,col_data):

        prop_name, prop_type, values = col_data
        if values is not None:
            #logging.debug("property '%s' has type '%s'" % ( prop_name, self.getADsTypeName(prop_type)))
            value=[]
            for v in values:
                conv=self.converters.get(prop_name, self._null_converter)
                value.append(conv(v[0]))
            if len(value) == 1:
                value = value[0]            
            return (prop_name, value)
        else:
            return (prop_name,None)
        
                          
    #def print_attribute(col_data):
        
        #prop_name, prop_type, values = col_data
        #if values is not None:
            #log(2, "property '%s' has type '%s'", prop_name, getADsTypeName(prop_type))
            #value = [converters.get(prop_name, _null_converter)(v[0]) for v in values]
            #if len(value) == 1:
                #value = value[0]
            #print " %s=%r" % (prop_name, value)
        #else:
            #print " %s is None" % (prop_name,)
    
    def search(self,searchFilter,attributes=None):
        gc = self.getGC()
        logging.debug("Search filter: `%s`, attributes: %s" %\
                      (searchFilter,str(attributes)))
        prefs = [(ADS_SEARCHPREF_SEARCH_SCOPE, (ADS_SCOPE_SUBTREE,))]
        hr, statuses = gc.SetSearchPreference(prefs)
        logging.debug("SetSearchPreference returned %d/%r" % ( hr, statuses))
        
        searchResults=[]
        h = gc.ExecuteSearch(searchFilter, attributes)
        try:
            
            hr = gc.GetNextRow(h)        
            while hr != S_ADS_NOMORE_ROWS:
                oneResult={}
                #print "-- new row --"
                if attributes is None:
                    # Loop over all columns returned
                    while 1:
                        col_name = gc.GetNextColumnName(h)
                        if col_name is None:
                            break
                        data = self._convert_attribute(gc.GetColumn(h, col_name))
                        
                        oneResult[data[0]]=data[1]
                else:
                    # loop over attributes specified.
                    for a in attributes:
                        try:
                            #data = gc.GetColumn(h, a)
                            data = self._convert_attribute(gc.GetColumn(h, a))
                            oneResult[data[0]]=data[1]
                        except adsi.error, details:
                            if details[0] != E_ADS_COLUMN_NOT_SET:
                                raise ADSISearcherException(str(details))
                            data=self._convert_attribute( (a, None, None) )
                            oneResult[data[0]]=data[1]
#                            oneResult.append(self._convert_attribute( (a, None, None) ))
                searchResults.append(oneResult)
                hr = gc.GetNextRow(h)
        except pywintypes.com_error,e:            
            if e.args[1]=="The size limit for this request was exceeded.":
                raise ADSISearcherException(e.args[0],code=ADSISearcherException.SIZE_EXCEEDED)
            raise e
        finally:
            gc.CloseSearchHandle(h)
        return searchResults
    
class UsefulADSISearches(ADSISearcher):
    

    # def getPrimaryGroupNameForUser(self,user):
    #     res=self.findUserOrGroup(user,['objectSid','primaryGroupID'])[0]
    #     objectSidPrefix=res['objectSid'].split('-')[0:-1]
    #     fullSid=objectSidPrefix+res['primaryGroupID']
    #     searchFilter="(objectSid=%s)" % fullSid
    #     return self.search(searchFilter,['cn','distinguishedName'])[0]

    def findGroups(self,groupName,attributes=None):
        searchFilter="(&(objectClass=group)(|(distinguishedName=%s)(name=%s)))" % (groupName,groupName)
        return self.search(searchFilter,attributes)
    
    def findUserOrGroup(self,objname,attributes=None):
        searchFilter="(&(|(objectClass=group)(objectClass=user))(|(distinguishedName=%s)(sAMAccountName=%s)))" % (objname,objname)
        return self.search(searchFilter,attributes)
    
    def findGroupMembers(self,groupName,objectClasses=None,attributes=None):
        #print "objectSid - enter function"
        # just same basic sanity check:
        assert re.search("\w+=\w+,",groupName),\
               "The group name must be in DN format. You supplied: %s" % groupName
        # get SID to get primaryGroupID
        # Q: why the f we are doing this? 
        # A: when we have a group that is set as someone's primaryGroupID and we want to
        #    get those users, we fetch the primaryGroupID here.
        try:
            searchFilter="(&(objectClass=group)(distinguishedName=%s))" % (groupName)
            objectSid=(self.search(searchFilter,["objectSid"])[0]['primaryGroupID'])
            primaryGroupID=objectSid.split["-"][:-1]
        except (KeyError,IndexError),e:
            logging.log(logging.DEBUG,"Couldn't find primaryGroupID")
            objectSid="DUMMY"
            primaryGroupID="-1"
        logging.log(logging.DEBUG,"objectSid: %s, RID: %s" % (objectSid,primaryGroupID))
        if type(objectClasses)==list and len(objectClasses)>0:
            searchFilter="(&(|(primaryGroupID=%s)(memberof=%s))(|%s))" % \
            (primaryGroupID,groupName,"".join(["(objectClass=%s)" % x for x in objectClasses]))
                        
        else:
            searchFilter="(memberof=%s)" % groupName        
        return self.search(searchFilter,attributes)
    
    
    
class TestADSISearcher(unittest.TestCase):
    def setUp(self):
        self.searcher=ADSISearcher()    
        
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
        
        
#def main():
    #global options
    #from optparse import OptionParser

    #parser = OptionParser()
    #parser.add_option("-f", "--file", dest="filename",
                      #help="write report to FILE", metavar="FILE")
    #parser.add_option("-v", "--verbose",
                      #action="count", default=1,
                      #help="increase verbosity of output")
    #parser.add_option("-q", "--quiet",
                      #action="store_true",
                      #help="suppress output messages")

    #parser.add_option("-U", "--user",
                      #help="specify the username used to connect")
    #parser.add_option("-P", "--password",
                      #help="specify the password used to connect")
    #parser.add_option("", "--filter",
                      #default = "(&(objectCategory=person)(objectClass=User))",
                      #help="specify the search filter")
    #parser.add_option("", "--attributes",
                      #help="comma sep'd list of attribute names to print")
    
    #options, args = parser.parse_args()
    #if options.quiet:
        #if options.verbose != 1:
            #parser.error("Can not use '--verbose' and '--quiet'")
        #options.verbose = 0

    #if args:
        #parser.error("You need not specify args")

    #search()

#if __name__=='__main__':
    #main()
