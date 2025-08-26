import datetime
import os
import tomllib
from logging import getLogger

logger = getLogger(__name__)

# Configuration handling class
class Config:
    class Redmine:
        def __init__(self):
            self.url          = None
            self.link_url     = None
            self.project_name = None
            self.login        = None
            self.username     = None
            self.password     = None

    class Filter:
        def __init__(self):
            self.sort             = None
            self.issue_id         = None
            self.query_id         = None
            self.parent_id        = None
            self.tracker_id       = None
            self.status_id        = None
            self.author_id        = None
            self.assigned_to_id   = None
            self.fixed_version_id = None

    def __init__(self):
        self._redmine = self.Redmine()
        self._filtter = self.Filter()

        self._font_name  = None
        self._tab_title  = None
        self._start_date = None
        self._end_date   = None
        self._holidays   = None

    def load_config_from_toml(self):
        """
        Load configuration from 'config.toml'.
        """

        config_file = 'config.toml' # constant file name
        if os.path.exists(config_file):
            with open(config_file, 'rb') as f:
                config = tomllib.load(f)

            redmine = config.get('redmine', {})
            account = redmine.get('account', {})
            filter = redmine.get('filter', {})
            spreadsheet = config.get('spreadsheet', {})
            gantt = spreadsheet.get('gantt', {})

            self._redmine.url = redmine.get('url', None).strip('/')
            self._redmine.link_url = self._redmine.url + '/issues/'
            self._redmine.project_name = redmine.get('project_name', None)

            self._redmine.login = account.get('need_login', False)
            self._redmine.username = account.get('username', None)
            self._redmine.password = account.get('password', None)

            self._filtter.sort = filter.get('sort', None)
            self._filtter.issue_id = filter.get('issue_id', None).replace(' ', '') if filter.get('issue_id', None) else None
            self._filtter.query_id = filter.get('query_id', None)
            self._filtter.tracker_id = filter.get('tracker_id', None)
            self._filtter.status_id = filter.get('status_id', None)
            self._filtter.assigned_to_id = filter.get('assigned_to_id', None)
            self._filtter.fixed_version_id = filter.get('fixed_version_id', None)

            self._font_name = spreadsheet.get('font_name', None)
            self._tab_title = spreadsheet.get('tab_title', None)

            self._start_date = gantt.get('start_date', None)
            self._end_date = gantt.get('end_date', None)

            self._holidays = config.get('holidays', [])
        else:
            logger.error(f"config file '{config_file}' not found.")
            return False

        # Validate mandatory fields
        if not all([self._redmine.url, self._redmine.project_name, 
                    self._start_date, self._end_date]):
            logger.error("Missing mandatory configuration fields.")
            return False

        self._start_date = datetime.datetime.strptime(self._start_date, '%Y/%m/%d').date()
        self._end_date = datetime.datetime.strptime(self._end_date, '%Y/%m/%d').date()

        return True

    def input_pw(self, prompt:str='Password: ') -> str:
        """
        Input password

        Args:
            prompt (str) : Input prompt string (='Password: ').

        Returns:
            (str) Entered string
        """

        from msvcrt import getch

        print(prompt, end='', flush=True)

        buf = ''
        n = 0
        while True:
            ch = getch()
            if ch == b'\r':
                print('')
                break
            elif ch == b'\x08':  # [BS]
                if n > 0:
                    buf = buf[:n-1]
                    n -= 1
                    print('\b \b', end='', flush=True)
            else:
                buf += ch.decode('UTF-8')
                print('*', end='', flush=True)
                n += 1

        return buf

    def user_account(self):
        if self._redmine.login:
            if not self._redmine.username:
                username = input('Username: ')
                self._redmine.username = username
            if not self._redmine.password:
                password = self.input_pw()
                self._redmine.password = password
        else:
            self._redmine.username = None
            self._redmine.password = None

    @property
    def url(self):
        return self._redmine.url

    @property
    def link_url(self):
        return self._redmine.link_url

    @property
    def project_name(self):
        return self._redmine.project_name

    @property
    def login(self):
        return self._redmine.login

    @property
    def username(self):
        return self._redmine.username

    @property
    def password(self):
        return self._redmine.password

    @property
    def sort(self):
        return self._filtter.sort

    @property
    def issue_id(self):
        return self._filtter.issue_id

    @property
    def query_id(self):
        return self._filtter.query_id

    @property
    def parent_id(self):
        return self._filtter.parent_id

    @property
    def tracker_id(self):
        return self._filtter.tracker_id

    @property
    def status_id(self):
        return self._filtter.status_id

    @property
    def author_id(self):
        return self._filtter.author_id

    @property
    def assigned_to_id(self):
        return self._filtter.assigned_to_id

    @property
    def fixed_version_id(self):
        return self._filtter.fixed_version_id

    @property
    def font_name(self):
        return self._font_name

    @property
    def tab_title(self):
        return self._tab_title if self._tab_title else self.project_name

    @property
    def start_date(self):
        return self._start_date

    @property
    def end_date(self):
        return self._end_date

    @property
    def holidays(self):
        return self._holidays

