from os import remove
from io import StringIO
from requests import get
from untangle import parse
from zipfile import ZipFile
from contextlib import closing
from os.path import isfile, dirname
from win32api import (
    GetFileVersionInfo,
    HIWORD, LOWORD,
)

class ChromeDriverInstaller():
    def __init__(self):
        if isfile('chromedriver.exe'):
            classFunctionType = type(self.autoInstall)
            for attribute in self.__dir__():
                if attribute.startswith('__') and attribute.endswith('__'):
                    continue
                if isinstance(self.__getattribute__(attribute), classFunctionType):
                    self.__dict__[attribute] = lambda: None

        drive = dirname(__file__)[:2]
        chrome_path = [
            fr'{drive}\Program Files\Google\Chrome\Application\chrome.exe',
            fr'{drive}\Program Files (x86)\Google\Chrome\Application\chrome.exe',
        ]
        chrome_path = list(
            filter(lambda chrome: isfile(chrome), chrome_path),
        )
        if not chrome_path:
            raise FileNotFoundError('Cannot find chrome!')

        def fetchFileVersion():
            info = GetFileVersionInfo(chrome_path[0], '\\')
            ms = info['FileVersionMS']
            ls = info['FileVersionLS']
            return [HIWORD(ms), LOWORD(ms), HIWORD(ls), LOWORD(ls)]

        self.fullVersion = fetchFileVersion()
        self.version = self.fullVersion[0]
        self.compatibleDriverVersion = None

    def getCompatibleDriverVersion(self):
        text = get('https://chromedriver.storage.googleapis.com/?delimiter=/&prefix=').text
        io = StringIO(text)
        xml = parse(io).ListBucketResult.CommonPrefixes
        compatible_driver_version = list(
            filter(lambda prefix: prefix.children[0].cdata.startswith(self.version.__str__()), xml),
        )
        self.compatibleDriverVersion = compatible_driver_version[0].children[0].cdata
        self.compatibleDriverVersion = self.compatibleDriverVersion[:-1] if self.compatibleDriverVersion.endswith('/') else self.compatibleDriverVersion
        return self.compatibleDriverVersion

    def installCompatibleDriver(self):
        url = f'https://chromedriver.storage.googleapis.com/{self.compatibleDriverVersion}/chromedriver_win32.zip'
        content = get(url).content
        with closing(open('chromedriver.zip', 'wb')) as f:
            f.write(content)
        
        with closing(ZipFile('chromedriver.zip', 'r', allowZip64=True)) as z:
            z.extract('chromedriver.exe')

        remove('chromedriver.zip')
        return None

    def autoInstall(self):
        self.getCompatibleDriverVersion()
        self.installCompatibleDriver()
        return None
