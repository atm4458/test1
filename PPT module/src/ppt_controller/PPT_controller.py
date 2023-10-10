# -*- coding: utf-8 -*-
"""
Created on Mon Oct  9 10:43:25 2023

@author: Jimmy
"""
import time
import win32com.client
import os
from .decorator import retry_protection
from . import template_parser as ppt_parse

file_template = os.path.abspath(r".\template\template.pptx")

location_template = os.path.abspath(r".\template\LocationTemplate")
class PPT_controller:
    def __init__(self, project_name, project_path = None):
        self.project_path = project_path or os.getcwd()
        self.project_name = project_name
        self.save_path = os.path.join(self.project_path, "AutoReport.pptx")
        self.slide=None
        self.slide_idx=None
    
    def iterate_slide(self, tmp_ppt):
        page_count = tmp_ppt.Slides.Count
        for page_idx in range(1, page_count+1):
            tmp_slide = tmp_ppt.Slides(page_idx)
            yield tmp_slide
    
    def iterate_shape(self, tmp_slide):
        shape_counts = tmp_slide.Shapes.Count
        for shape_idx in range(1, shape_counts+1):
            tmp_shape = tmp_slide.Shapes(shape_idx)
            yield tmp_shape
    
    def ReadTemplate(self):
        template_files = [file for file in os.listdir(location_template)
                            if file.endswith(".pptx")]
        location_information = {}
        for get_file in template_files:
            file_type = get_file.replace(".pptx", "")
            full_path = os.path.join(location_template, get_file)
            try:
                tmp_ppt = self.app.Presentations.Open(full_path)
                for tmp_slide in self.iterate_slide(tmp_ppt):
                    template_location = []
                    for tmp_shape in self.iterate_shape(tmp_slide):
                        loc_params = ppt_parse.acquire_template_info(tmp_shape)
                        template_location.append(loc_params)
                location_information[file_type] = template_location
            except Exception as e:
                raise ValueError(e)
            finally:
                tmp_ppt.Close()
        return location_information
    
    @retry_protection
    def Initialize(self):
        self.time_stamp = time.time()
        self.app = win32com.client.Dispatch('PowerPoint.Application')
        self.prs = self.app.Presentations.Open(file_template)
        self.slide_idx = 2
        self.Save(self.save_path)
    
    @retry_protection
    def Close(self):
        self.prs.Close()
    
    @retry_protection
    def Save(self, save_as_path = None):
        self.save_path = save_as_path or self.save_path
        self.prs.SaveAs(self.save_path)
    
    @retry_protection
    def CreateSlide(self, title=None):
        self.prs.Slides.Add(self.slide_idx, 12)
        self.slide = self.prs.Slides(self.slide_idx)
        self.slide_idx+=1
    
    def AddText(self, text_string, **kwargs):
        Left, Top, Width, Height= [kwargs.get(get_loc, None) or 100 
                                   for get_loc in ["LEFT", "TOP", "WIDTH", "HEIGHT"]]
        text_box = self.slide.Shapes.AddTextbox(Orientation=0x1,
                                                Left=Left,Top=Top,Width=Width,Height=Height)
        text_box.TextFrame.TextRange.Text = text_string
    
    def AddPicture(self, image_path, **kwargs):
        shape = self.slide.Shapes.AddPicture(FileName=image_path, LinkToFile=False, 
                                             SaveWithDocument=True, Left=0, Top=0)
        