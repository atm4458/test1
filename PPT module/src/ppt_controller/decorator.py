# -*- coding: utf-8 -*-
"""
Created on Mon Oct  9 11:53:04 2023

@author: Jimmy
"""
import os, time
from functools import wraps
from pythoncom import com_error

def retry_protection(func, *arg, **args):
    @wraps(func)
    def wrap_method(*arg, **args):
        success = False
        e=None
        for retry in range(4):
            try:
                result = func(*arg, **args)
                print(retry)
                success = True
                break
            except AttributeError as e:
                print(f"請注意是否影響到ppt運行。...{e}")
            except com_error as e:
                if "這份簡報是唯讀檔案，必須換用檔名儲存" in e.args[2][2]:
                    save_path = os.path.join(arg[0].project_path, "AutoReport_new.pptx")
                    func(*arg, save_path)
            time.sleep(1)
            print(f"動作失敗...重試中...")
        if not success:
            try:
                result = func(*arg, **args)
            except Exception as e:
                raise ValueError(e)
        return result
    return wrap_method