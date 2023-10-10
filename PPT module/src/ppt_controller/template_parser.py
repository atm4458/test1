# -*- coding: utf-8 -*-
"""
Created on Mon Oct  9 14:56:33 2023

@author: Jimmy
"""


def acquire_template_info(tmp_shape):
    content = tmp_shape.TextFrame.TextRange.Text.split("\r")
    name = content[0]
    args = {}
    for arg in content[1:]:
        if "=" in arg:
            key, val = arg.split("=")
            args[key.strip()] = val.strip()
    
    position = {"Left":tmp_shape.Left,
                "Top":tmp_shape.Top,
                "Width":tmp_shape.Width,
                "Height":tmp_shape.Height}
    loc_params = [name, args, position]
    return loc_params