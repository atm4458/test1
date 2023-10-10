# -*- coding: utf-8 -*-
"""
Created on Mon Oct  9 10:43:09 2023

@author: Jimmy
"""

from src import PPT_controller

ppt = PPT_controller(project_name="test")
ppt.Initialize()
location_information = ppt.ReadTemplate()
ppt.Save()
ppt.CreateSlide()

ppt.AddText("Hello")

ppt.AddPicture(r"pngtree-cute-cartoon-light-bulb-image_1134759.jpg")

ppt.Close()
