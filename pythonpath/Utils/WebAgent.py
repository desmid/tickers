###########################################################################
import logging
Logger = logging.getLogger('LoadPrices')
Logger.debug("Load: Utils.WebAgent")

###########################################################################
import sys
try:
    #Python3
    from urllib.request import Request, urlopen
    from urllib.error import URLError, HTTPError
    from urllib.parse import urlencode
except:
    #Python2
    from urllib2 import Request, urlopen, URLError, HTTPError
    from urllib import urlencode

###########################################################################
class WebAgent(object):
    """
    Class to access web pages. Wraps Python3 or Python2 implementation details.
    Keeps a record of the request:

    (1) requested url, (2) real url retrieved, (3) HTTP response code,
    (4) returned header info, (5) number of tries, (6) maximum tries allowed,
    (7) last request timeout, (8) last exception message, (9) the page itself.

    Constructor and usage:

    webAgent = WebAgent()

    html = webAgent.fetch(url)

    if webAgent.ok():
        print(html)
    else:
        print( "Diagnostics:\n" + str(webAgent) )
        if webAgent.response_code() == 404:
           #do something with this situation
        ...
        
    Public methods:

    fetch(url) : fetches and returns the web page:
      returns 'no response' if URL cannot be retrieved after preset retries
      and timeouts, or is invalid.

    ok()              returns True/False as fetch succeeded/failed

    url()             returns original URL
    real_url()        returns real URL retrieved
    response_code()   returns HTTP response code as integer (200, 404, etc.)
    error()           returns exception/error condition
    info()            returns server headers as a dict (Content-Type, etc.)
    html()            returns already fetched web page

    See: https://docs.python.org/3.4/howto/urllib2.html  (Python3)
         https://docs.python.org/2/howto/urllib2.html    (Python2)
    """

    Header = {
        'User-Agent': 'Mozilla/5.0 AppleWebKit/537.11 (KHTML, like Gecko) Chrome/23.0.1271.64 Safari/537.11',
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
        'Accept-Charset': 'ISO-8859-1,utf-8;q=0.7,*;q=0.3',
        'Accept-Encoding': 'none',
        'Accept-Language': 'en-US,en;q=0.8',
        'Connection': 'keep-alive',
    }

    Deft_Html    = 'no response'
    Deft_Timeout = 10
    MaxTries     = 5
    NoError      = None

    def __init__(self, paramDict=None):
        self.params = paramDict
        self.state = {
            'url':     None,  #supplied url
            'realurl': None,  #actual url retrieved (possible redirect)
            'status':  None,  #http response code (200, 404, etc)
            'error':   None,  #exception raised with error message of last try
            'tries':   None,  #number of tries
            'timeout': None,  #timeout of last try
            'info':    None,  #meta info
            'html':    None,  #the retrieved page
        }
        self._reset_state()
        
    def _reset_state(self, url=None):

        def get_timeout():
            try:
                t = int(self.params['webTimeOut'])
                if t < 1:
                    t = self.Deft_Timeout
            except Exception:
                t = self.Deft_Timeout
            return t

        self.state['url']     = url
        self.state['realurl'] = None
        self.state['status']  = None
        self.state['error']   = self.NoError
        self.state['tries']   = 0
        self.state['timeout'] = get_timeout()
        self.state['info']    = {}
        self.state['html']    = self.Deft_Html

    def fetch(self, url):
        if not isinstance(url, str):
            self.state['error'] = "URL must be a string '{}'".format(str(url))
            return self.state['html']

        self._reset_state(url)

        self.state['tries'] = 1

        while True:
            Logger.debug('try {!s}/{!s}/{!s}'.format(self.state['tries'],
                                                     self.MaxTries,
                                                     self.state['timeout']))

            try:
                req = Request(url, headers=self.Header)
                response = urlopen(req, None, self.state['timeout'])  #with timeout
            except HTTPError as e:
                self.state['error'] = 'HTTPError: ' + str(e.code)
            except URLError as e:
                self.state['error'] = 'URLError: ' + str(e.reason)
            except Exception as e:
                self.state['error'] = 'Exception: ' + str(e)
            else:
                self.state['status'] = response.getcode()
                self.state['realurl'] = response.geturl()
                self.state['info'] = response.info()
                self.state['html'] = response.read().decode('utf-8', 'ignore')
                self.state['error'] = self.NoError  #cleanup
                break  #got something

            Logger.error(self.state['error'])

            if self.state['tries'] >= self.MaxTries:
                break  #give up

            self.state['timeout'] *= 2
            self.state['tries'] += 1

        return self.state['html']

    #pretty-print the object diagnostics for print() and str() calls
    def __str__(self):
        triesmaxtime = '{!s}/{!s}/{!s}'.format(
            self.state['tries'], self.MaxTries, self.state['timeout'])
        s = ("  {:<}: {!s}" * 6)[2:]
        s = s.format(
            'status',  self.state['status'],
            'tries/max/timeout', triesmaxtime,
            'error',   self.state['error'],
            'url',     self.state['url'],
            'realurl', self.state['realurl'],
            'info',    sorted(self.state['info'].items()),
        )
        return s

    def ok(self):
        #need redirect?
        if self.state['status'] == 200: return True
        if self.state['html'] != self.Deft_Html: return True
        if self.state['error'] == self.NoError: return True
        return False

    def html(self):          return self.state['html']
    def response_code(self): return self.state['status']
    def url(self):           return self.state['url']
    def real_url(self):      return self.state['realurl']
    def error(self):         return self.state['error']
    def info(self):          return self.state['info']

###########################################################################
if __name__ == '__main__':
    def make_url(url, values):
        if url is not None and values is not None:
            return url + '?' + urlencode(values)
        return url

    o = WebAgent()

    url = 1234
    print('')
    print('Trying URL: ' + str(url))
    o.fetch(url)
    if o.ok():
        print("Fetch OK")
    else:
        print("Fetch FAILED")
    print(o)
    print(o.html())

    url = 'this is garbage'
    print('')
    print('Trying URL: ' + url)
    o.fetch(url)
    if o.ok():
        print("Fetch OK")
    else:
        print("Fetch FAILED")
    print(o)
    print(o.html())

    url = make_url('https://valid.looking.url', [('d',4), ('e',5), ('f',6)])
    print('')
    print('Trying URL: ' + url)
    o.fetch(url)
    if o.ok():
        print("Fetch OK")
    else:
        print("Fetch FAILED")
    print(o)
    print(o.html())

    url = 'https://query1.finance.yahoo.com/v7/finance/quote?symbols=^FTSE,^FTAS'
    print('')
    print('Trying URL: ' + url)
    o.fetch(url)
    if o.ok():
        print("Fetch OK")
    else:
        print("Fetch FAILED")
    print(o)
    print(o.html())

    print('End')

###########################################################################
