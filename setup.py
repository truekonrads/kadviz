# Author: Konrads Klints <konrads.smelkovs@kpmg.co.uk>
from distutils.core import setup
import py2exe
import sys
# ...
# ModuleFinder can't handle runtime changes to __path__, but win32com uses them
try:
    # py2exe 0.6.4 introduced a replacement modulefinder.
    # This means we have to add package paths there, not to the built-in
    # one.  If this new modulefinder gets integrated into Python, then
    # we might be able to revert this some day.
    # if this doesn't work, try import modulefinder
    try:
        import py2exe.mf as modulefinder
    except ImportError:
        import modulefinder
    import win32com
    for p in win32com.__path__[1:]:
        modulefinder.AddPackagePath("win32com", p)
    for extra in ["win32com.shell"]:  # ,"win32com.mapi"
        __import__(extra)
        m = sys.modules[extra]
        for p in m.__path__[1:]:
            modulefinder.AddPackagePath(extra, p)
except ImportError, e:
    # no build path setup, no worries.
    print str(e)
    pass
setup(console=['kadviz.py', 'plotrepl.py', 'trusts.py'],
      options={"py2exe":
               {
                   "includes": ["win32com", "win32com.adsi", "pydot", "ldap"],
                   "dll_excludes": ["MSVCP90.dll", "libzmq.pyd", "geos_c.dll", "api-ms-win-core-string-l1-1-0.dll", "api-ms-win-core-registry-l1-1-0.dll", "api-ms-win-core-errorhandling-l1-1-1.dll", "api-ms-win-core-string-l2-1-0.dll", "api-ms-win-core-profile-l1-1-0.dll", "api-ms-win*.dll", "api-ms-win-core-processthreads-l1-1-2.dll", "api-ms-win-core-libraryloader-l1-2-1.dll", "api-ms-win-core-file-l1-2-1.dll", "api-ms-win-security-base-l1-2-0.dll", "api-ms-win-eventing-provider-l1-1-0.dll", "api-ms-win-core-heap-l2-1-0.dll", "api-ms-win-core-libraryloader-l1-2-0.dll", "api-ms-win-core-localization-l1-2-1.dll", "api-ms-win-core-sysinfo-l1-2-1.dll", "api-ms-win-core-synch-l1-2-0.dll", "api-ms-win-core-heap-l1-2-0.dll", "api-ms-win-core-handle-l1-1-0.dll", "api-ms-win-core-io-l1-1-1.dll", "api-ms-win-core-com-l1-1-1.dll", "api-ms-win-core-memory-l1-1-2.dll", "api-ms-win-core-version-l1-1-1.dll", 
                   "api-ms-win-core-version-l1-1-0.dll",
                   "api-ms-win-core-delayload-l1-1-0.dll",
                   'api-ms-win-core-delayload-l1-1-1.dll',
                   'api-ms-win-core-delayload-l1-1-1.dll',
                   'api-ms-win-core-localization-l1-2-0.dll',
                   'api-ms-win-core-sysinfo-l1-1-0.dll',
                   'api-ms-win-core-errorhandling-l1-1-0.dll',
                    'api-ms-win-core-file-l1-1-0.dll',
                    'api-ms-win-core-timezone-l1-1-0.dll',
                    'api-ms-win-core-processenvironment-l1-1-0.dll',
                    'api-ms-win-core-rtlsupport-l1-1-0.dll',
                    'api-ms-win-security-base-l1-1-0.dll',
                    'api-ms-win-core-localization-obsolete-l1-2-0.dll',
                    'api-ms-win-core-string-obsolete-l1-1-0.dll',
                    'api-ms-win-crt-private-l1-1-0.dll',
                   'api-ms-win-core-processthreads-l1-1-0.dll',
                    'api-ms-win-core-processthreads-l1-1-1.dll',
                    'api-ms-win-crt-string-l1-1-0.dll',
                    'api-ms-win-crt-runtime-l1-1-0.dll',
                     'api-ms-win-core-heap-l1-1-0.dll',
                     'api-ms-win-core-interlocked-l1-1-0.dll',
                     'api-ms-win-core-debug-l1-1-0.dll',
                   "api-ms-win-core-synch-l1-1-0.dll"]

               }
               })
# http://www.py2exe.org/index.cgi/UsingEnsureDispatch
