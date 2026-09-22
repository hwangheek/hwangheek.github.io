import sys
import platform
assert sys.platform == 'win32', 'xingAPI는 Windows 환경에서 사용 가능합니다.'
assert platform.architecture()[0] == '32bit', 'xingAPI는 32bit 환경에서 사용 가능합니다.'

import os
import pandas as pd
import pytz
import re
import time
from datetime import date, datetime, timedelta
from getpass import getpass
from pythoncom import PumpWaitingMessages
from win32com.client import DispatchWithEvents

import logging
logger = logging.getLogger(__name__)
logger.setLevel(logging.DEBUG)
_h = logging.StreamHandler()
_h.setLevel(logging.DEBUG)
logger.addHandler(_h)

XINGAPI_PATH = '/eBEST/xingAPI/'


def build_meta_res():
    """ res 파일들의 meta data
        
        Example
        -------
        >>> build_meta_res()
        {
            't8413': {
                'desc': '주식챠트(일주월)',
                'input': {
                    't8413InBlock': {
                        'occurs': False,
                        'fields': [
                            {
                                'name': 'shcode',
                                'desc': '단축코드',
                                'type': 'char',
                                'size': 6
                            },
                            { ... },
                            ...
                        ]
                    }
                },
                'output': {
                    't8413OutBlock1': {
                        'occurs': True,
                        'fields': [ 'price', ... ]
                    },
                    ...
                }
            },
            ...
        }
    """
    meta = {}
    
    fnames = filter(
        lambda x: not re.match(r'.*\_\d+\.res$', x),
        os.listdir(os.path.join(XINGAPI_PATH, 'res'))
    )
    
    def parse_field(line):
        cols = line.split(',')
        return {
            'name': cols[1].strip(),
            'desc': cols[0].strip(),
            'type': cols[3].strip(),
            'size': cols[4].strip()
        }
    
    def parse_file(lines):
        parsed = {}
        lines = list(map(lambda x: x.replace('\t','').replace('\n','').replace(';','').strip(), lines))
        lines = list(filter(lambda x:x, lines))
        for i in range(len(lines)):
            if '.Func' in lines[i] or '.Feed' in lines[i]:
                parsed['desc'] = lines[i].split(',')[1].strip()
            elif lines[i] == 'begin':
                latest_begin = i
            elif lines[i] == 'end':
                block_info = lines[latest_begin-1].split(',')
                
                if not block_info[2] in parsed:
                    parsed[block_info[2]] = {}
                
                parsed[block_info[2]][block_info[0]] = {
                    'occurs': 'occurs' in block_info,
                    'fields': list(map(parse_field, lines[latest_begin+1:i]))
                }
        return parsed
    
    for fname in fnames:
        meta[fname.replace('.res','')] = parse_file(
            open(os.path.join(XINGAPI_PATH, 'res/', fname)).readlines()
        )
    
    return meta


meta_res = build_meta_res()


class _SessionHandler:
    def OnLogin(self, code, msg):
        """ 서버와의 로그인이 끝나면 실행되는 함수

            @arg code[str] 서버에서 받은 메시지 코드
            @arg msg[str] 서버에서 받은 메시지 정보
        """
        self.waiting = False
    
        if code == '0000':
            logger.info('[*] 로그인 성공')
        else:
            logger.warning('[*] 로그인 실패 : {}'.format(msg))
    
    def OnDisconnect(self):
        """ 서버와의 연결이 끊어졌을 때 실행되는 함수
        """
        self.waiting = False
        
        logger.info('[*] 서버와의 연결이 끊어졌습니다')

_session = DispatchWithEvents('XA_Session.XASession', _SessionHandler)

def login(
    server=None,
    username=None,
    password=None,
):
    """ 로그인
    """
    # 기존에 연결되어 있는 서버가 있으면, 연결을 끊는다
    if _session.IsConnected():
        _session.DisconnectServer()
    
    # 로그인 시 필요한 정보를 입력받는다
    login_server = (server or input('[*] 접속 서버 ((r)eal / (D)emo / (a)ce) : ')).lower()[:1]
    login_server = {
        'r': 'hts.ebestsec.co.kr',
        'd': 'demo.ebestsec.co.kr',
        'a': '127.0.0.1'
    }.get(login_server, 'demo.ebestsec.co.kr')
    login_port = 20001
    login_id = username or input('[*] 아이디 : ')
    login_pw = password or getpass('[*] 패스워드 : ')
    login_cert = '' if login_server == 'demo.ebestsec.co.kr' else getpass('[*] 공인인증서 암호 : ')
    
    # 로그인 요청을 보낸다
    _session.waiting = True
    _session.ConnectServer(login_server, login_port)
    _session.Login(login_id, login_pw, login_cert, 0, 0)
    while _session.waiting:
        PumpWaitingMessages()
        time.sleep(0.05)

def accounts():
    """ 계좌 리스트
    """
    accounts = []
    
    num_account = _session.GetAccountListCount()
    for i in range(num_account):
        acc_no = _session.GetAccountList(i)
        accounts.append({
            'acc' : acc_no,
            'nm' : _session.GetAccountName(acc_no),
            'detail' : _session.GetAcctDetailName(acc_no),
            'nick' : _session.GetAcctNickname(acc_no)
        })
    
    return accounts

def account(index=0):
    """ 계좌번호

        @arg index[*int=0] 불러올 계좌의 순번
    """
    return _session.GetAccountList(index)



""" Query
"""
_query_status = {}

class _QueryHandler:
    def __init__(self):
        self.response = {}
        self.decomp = False
        self.qrycnt = None
        self.waiting = False
        self.res = None
    
    def init(self, res):
        self.LoadFromResFile('/Res/{}.res'.format(res))
        self.res = res
    
    def set_data(self, block, k, v, index=0):
        if k == 'comp_yn' and v.lower() == 'y':
            self.decomp = True
        elif k == 'qrycnt':
            self.qrycnt = int(v)
        
        self.SetFieldData(block, k, index, v)
    
    def get_block_data(self, block, index=0):
        block_data = {}
        for field in meta_res[self.res]['output'][block]['fields']:
            data = self.GetFieldData(block, field['name'], index)
            
            if field['type'] == 'long':
                if data == '-':
                    data = 0
                data = int(data or 0)
            elif field['type'] == 'double' or field['type'] == 'float':
                data = float(data or 0.0)
            
            block_data[field['name']] = data
        
        return block_data
    
    def OnReceiveData(self, res):
        """ 요청 데이터 도착 Listener
            
            self.GetFieldData(...)를 통해 전송받은 데이터 확인이 가능하다.

            @arg res[str] 요청 res 파일명
        """
        # decompress가 필요한 경우 압축해제
        # TODO : OutBlock1 말고 다른 occurs가 있는 케이스 (ex. FOCCQ33600)
        if self.decomp:
            self.Decompress(res + 'OutBlock1')
        
        for block in meta_res[res]['output'].keys():
            # 해당 블럭이 occurs인 경우,
            if meta_res[res]['output'][block]['occurs']:
                row_res = []
                for i in range(self.GetBlockCount(block)):
                    row_res.append(self.get_block_data(block, i))
            # 해당 블럭이 단일 데이터인 경우,
            else:
                row_res = self.get_block_data(block)
        
            self.response[block] = row_res
        
        self.waiting = False

def query(res, send, cont=False, timeout=10):
    """ Query 요청

        @arg res[str]`t1102` 사용할 res 파일명
        @arg send[dict] 전송할 데이터
            {
                'Block1': [{'Field1': 'Value1', 'Field2': 'Value2'}, {...}, {...}],
                'Block2': {'Field3': 'Value3', 'Field4': 'Value4'}
            }
    
            단일 InBlock의 경우에는 아래와 같이 간단한 형식도 입력받음
            {'Field1': 'Value1', 'Field2': 'Value2'}
        @arg cont[*bool=False] 연속조회 여부
        @arg timeout[*int=10] 서버 응답 최대 대기 시간, -1인 경우 infinite time
    """
    # res 파일 로드
    _query = DispatchWithEvents('XA_DataSet.XAQuery', _QueryHandler)
    _query.init(res)
    
    if not cont:
        # 전송 현황 업데이트
        if not res in _query_status:
            _query_status[res] = []
        
        while _query_status[res] and _query_status[res][-1] + 1 < time.time():
            _query_status[res].pop()
        
        # 초당 전송 횟수를 고려
        tr_count_per_sec = _query.GetTRCountPerSec(res)
        if len(_query_status[res]) >= tr_count_per_sec:
            delay = max(_query_status[res][-1] + 1.05 - time.time(), 0)
            time.sleep(delay)
        
        # 기간(10분)당 전송 횟수를 고려
        # TODO : 10분 제한이 걸리면 blocking state 진입
        tr_count_limit = _query.GetTRCountLimit(res)
        while tr_count_limit and _query.GetTRCountRequest(res) >= tr_count_limit:
            time.sleep(1)
            _query = DispatchWithEvents('XA_DataSet.XAQuery', _QueryHandler)
            _query.init(res)
    
    # simplified 된 input를 받았을 경우
    send_first_value = list(send.values())[0]
    if not (
        isinstance (send_first_value, list) or
        isinstance (send_first_value, dict)
    ):
        send = { '{}InBlock'.format(res): send }
    
    # 전송할 데이터를 설정
    for block in send.keys():
        if isinstance(send[block], dict):
            for (k, v) in send[block].items():
                _query.set_data(block, k, v)
        elif isinstance(send[block], list):
            for i in range(len(send[block])):
                for (k, v) in send[block][i].items():
                    _query.set_data(block, k, v, i)
        else:
            raise ValueError('알 수 없는 형태의 데이터입니다')
    
    else:
        time.sleep(0.05)
    
    # 데이터 요청
    _query.Request(cont)
    
    now = time.time()
    if not cont:
        _query_status[res].insert(0, now)
    _query.waiting = True
    while _query.waiting:
        if timeout >= 0 and now + timeout < time.time():
            _query.waiting = False
            raise TimeoutError('Query Timeout')
        PumpWaitingMessages()
    
    return _query.response

class _RealtimeHandler:
    def OnReceiveRealData(self, res):
        response = {}
        for field in meta_res[res]['output']['OutBlock']['fields']:
            response[field['name']] = self.GetFieldData('OutBlock', field['name'])

        self.callback(res, response)

class Realtime:
    def __init__(self, res, callback):
        self._res = res
        self._instance = DispatchWithEvents('XA_DataSet.XAReal', _RealtimeHandler)
        self._instance.LoadFromResFile(f'/Res/{res}.res')
        self._instance.callback = callback
        
        self.subscribed_keys = []
    
    def subscribe(self, key=None):
        if key in self.subscribed_keys:
            print(f'{self._res}는 이미 {key} 데이터를 수신 중입니다.')
            return None
        
        if key:
            self._instance.SetFieldData('InBlock', meta_res[self._res]['input']['InBlock']['fields'][0]['name'], key)
        self._instance.AdviseRealData()
        
        self.subscribed_keys.append(key)
    
    def unsubscirbe(self, key=None):
        if key is None:
            self._instance.UnadviseRealData()
        else:
            if key not in self.subscribed_keys:
                raise ValueError(f'{self._res}는 {key} 데이터를 수신하고 있지 않습니다.')
            self._instnace.UnadviseRealDataWithKey(key)
    
    @staticmethod
    def listen(delay=.01):
        while True:
            PumpWaitingMessages()
            time.sleep(delay)
