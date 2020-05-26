import traceback
import numpy as np
import cv2,random,os,sys
from time import sleep as tmSleep
import logging
import time
from moviepy.audio.io.AudioFileClip  import AudioFileClip
import xlrd

from openpyxl import load_workbook
from datetime import datetime

from moviepy.video.io.VideoFileClip import VideoFileClip
from moviepy.video.VideoClip import ImageClip
from moviepy.video.compositing.CompositeVideoClip import CompositeVideoClip
from moviepy.video.compositing.concatenate import concatenate_videoclips     # new addition 
from moviepy.audio.AudioClip import CompositeAudioClip


from PyQt5 import QtWidgets,QtGui,QtCore
from PyQt5.QtWidgets import * #QApplication,QToolTip,QLabel, QMainWindow ,QWidget,QDesktopWidget,QMessageBox,QLineEdit,QGridLayout, QFileDialog ,QAction, qApp,QMenu,QFrame,QColorDialog 
from PyQt5.QtGui import * #QIcon, QColor,QIntValidator
from PyQt5.QtCore import * #Qt,QRunnable,QObject, QThread,QThreadPool, pyqtSignal, pyqtSlot
import sys
from PIL import ImageColor,Image, ImageFont, ImageDraw 
from pathlib import Path
import shutil






class Worker(QThread):
    countChanged = pyqtSignal(int)  # progress bar
    start_activation= pyqtSignal(bool) # start Button Activation deActivation
    stopwork=pyqtSignal(bool)# pause the process
    setError=pyqtSignal(str) # Send Error to GUI
    try:
        log_file_path=""
        temp_path=""
        Images_files_path_collection=[]
        new_audfile=[]
        filepath=""
        stop_process=False
        start_process=False
        Limit_aud=0;
        Intro_arr=[] 
        Outro_arr=[]
        intro_sel_path=""
        outro_sel_path=""
    except Exception as error:
        msg=str(error)
        #self.setError.emit(msg)

 
    def __init__(self,Excel_filepath,Image_folder_path,Input_intro_path,Input_outro_path,Input_audio_path,Input_dest_path,Input_text_color,Input_back_color,Input_time_Perslide,Aud,No_of_Row):
        
        QThread.__init__(self)
        self.Excel_filepath = Excel_filepath  #GLobal class Variavbles
        self.Image_folder_path=Image_folder_path
        self.Input_intro_path=Input_intro_path
        self.Input_outro_path=Input_outro_path
        self.Input_audio_path=Input_audio_path
        self.Input_text_color=Input_text_color
        self.Input_back_color=Input_back_color
        self.Input_time_Perslide=Input_time_Perslide
        self.Input_dest_path=Input_dest_path
        self.Aud=Aud;
        self.No_of_Row=No_of_Row
        temp_path=self.temp_path
        Images_files_path_collection=self.Images_files_path_collection
        filepath=self.filepath
        stop_process=self.stop_process
        log_file_path=self.log_file_path
        new_audfile=self.new_audfile
        Limit_aud=self.Limit_aud
        Intro_arr=self.Intro_arr 
        Outro_arr=self.Outro_arr
        intro_sel_path=self.intro_sel_path
        outro_sel_path=self.outro_sel_path
        
        
    #@pyqtSlot()
    def on_stopprocess(self,val):
        try:
            if(val==True):
                self.stop_process=True
                print("QThread terminated")
                print('IN THREAD AREA WE R',self.stop_process)
                self.Logger("process Stop Successfull")
        except Exception as error:
            msg=str(error)
            var = traceback.format_exc()
            print(var)
            self.Logger(var)
            self.setError.emit(msg)

    def run(self):
        try:
            if self.stop_process==False:
                print("process is going to stop")
                start_btn=False
                self.start_activation.emit(start_btn)
                
                self.store_audio()  # Store Aud Files
                self.store_intro_outro_path()   # store Intro outro Paths
                print("Intro Arr",self.Intro_arr)
                print("Outro Arr",self.Outro_arr)
                TIME_LIMIT=100
                self.Create_temp_directory();
                self.resize_images()
                print("OK")
                Row_Of_Questions=[];
                temp_arr=[]
                wb = xlrd.open_workbook(self.Excel_filepath) 
                sheet = wb.sheet_by_index(0) 

                arr=[]
                rows = sheet.nrows
                columns = sheet.ncols
                self.No_of_Row=columns
                print("rows:",rows)
                print("cols:",columns)

                for k in range(sheet.nrows):
                    temp_arr.append(sheet.row_values(k))
                
                row_count=0;
                for val in temp_arr: 
                    if val != None : 
                        Row_Of_Questions.append(val)
                        row_count=row_count+1;
                #print(Row_Of_Questions)
                
        
        
                count=0;
                count_add_audio=0;
                video_count=0;
                video_path=[];
                count_progress=0;

                print("No of Rows:",self.No_of_Row)
                print(TIME_LIMIT)
                p_dur=0
            
                for x in Row_Of_Questions:
                    print(x)
                    if self.stop_process == False:
                        p_dur=TIME_LIMIT/row_count # for progressBar
                        print("No _of Rows",row_count)
                        print("count_progress:",count_progress)
                        print("p_dur:",p_dur)
                        count=0;
                        #questions=x.split(',');
                        img_paths=[]
                        count_for_image_path=0;
                        for q in x:
                            if q:
                                if self.stop_process == False:
                                    #print("Questions at a Row :",q)
                                    count=count+1;
                                    count_for_image_path=count_for_image_path+1;
                                    #print("NextRow")
                                    path=self.create_image(q,count,count_for_image_path);
                                    img_paths.append(path);
                        video_count=video_count+1;
                        if self.stop_process == False:
                            self.create_video(img_paths,video_count);
                        if self.stop_process == False:
                            count_progress =count_progress+p_dur
                            if(count_progress>0 and count_progress<=TIME_LIMIT):
                                self.countChanged.emit(count_progress)
                        #count_progress =count_progress+p_dur
                        #self.Add_Audio(video_path,video_count);
                        #QApplication.processEvents()
                self.Remove_temp_files()
                start_btn=True
                self.start_activation.emit(start_btn)
                self.Logger("Excel file Row fetch part Done Success")
        except Exception as error:
            msg=str(error)
            var = traceback.format_exc()
            #print(var)
            self.Logger(var)
            self.setError.emit(msg)
            

    def Remove_temp_files(self):
        try:
            dir_path = self.temp_path
            #print("deleted files:",dir_path)
            try:
                shutil.rmtree(dir_path)
                self.Logger("Temporary files Deleted Permanetly Sucess")
            except OSError as error:
                #print("Error: %s : %s" % (dir_path, error.strerror))
                msg=str(error)
                self.setError.emit(msg)
        except Exception as error:
            var = traceback.format_exc()
            self.Logger(var)
            msg=str(error)
            self.setError.emit(msg)
        
    def Create_temp_directory(self):
        try:
            d=chr(92)
            self.temp_path=self.Input_dest_path+d+"Temp"
            Path(self.temp_path).mkdir(parents=True, exist_ok=True) #  creating a new directory
          #  print(self.temp_path)
            self.Logger("Temporaray directory created Successfully")
        except Exception as error:
            msg=str(error)
            var = traceback.format_exc()
            self.Logger(var)
            self.setError.emit(msg)
        
    def resize_images(self):
        try:
            os.chdir(self.Image_folder_path)  
            path=self.Image_folder_path   #self.Image_path_first
            for file in os.listdir('.'): 
                if file.endswith(".jpg") or file.endswith(".jpeg") or file.endswith("png"): 
                    # opening image using PIL Image 
                    im = Image.open(os.path.join(path, file))     
                  # im.size includes the height and width of image 
                    width, height = im.size    
                  #  print(width, height) 
                    self.Images_files_path_collection.append(im.filename)
                    self.Logger("Find Images from Input Folder Sucessfully")
        except Exception as error:
            msg=str(error)
            var = traceback.format_exc()
            self.Logger(var)
            self.setError.emit(msg)
    
    def Check_Caps(self,test_text):
        try:
            word=test_text
            count=0;
            if word.islower():
                count=1;
            if word.isupper():
                count=2
            if not word.islower() and not word.isupper():
                count=1
            self.Logger("Check Capital Letters Sucessfully")
            return count;
        except Exception as e:
            var = traceback.format_exc()
            self.Logger(var)
  
    def create_image(self,text,i,count_for_image_path):
        try:
        
            frame_no=i;
            count_img=count_for_image_path;
            test_text=text
            
            path=random.choice(self.Images_files_path_collection)  
                
            image=cv2.imread(path);
            image = cv2.resize(image,(1280,720))
            overlay = image.copy()
            image_new=""
            #image = cv2.resize(image,(1280,720))  # resize the image 
##            font = cv2.FONT_HERSHEY_COMPLEX; 
##            org = (70, 80)
##            org_center=(400,80)
##            org_large_center=(200,80)
            fontScale = 1
            print("color:",self.Input_text_color)
            color = self.Input_text_color #(255, 255, 255, .4) 
            thickness = 2  # Line thickness of 2 px
            alpha = 0.15
            color_transparent = self.Input_back_color#(255, 20, 147);
            print("Text COlor __________",color)
            print("Tranparent COlor __________",color_transparent)
                
       #     print("length_of_input_string:",len(test_text))
            #print(image.shape)
        
        
            labelSize=cv2.getTextSize(test_text,cv2.FONT_HERSHEY_COMPLEX,1,1)
            width_of_rectangle=labelSize[0][0]+20;
         #   print("Text_size_width",labelSize[0][0])
            
            #ori=cv2.rectangle(image, (50, 30), (1230,100), (255, 255, 255, .4), 10)
            val=0;
            val=self.Check_Caps(test_text)
            rectangle_count=0;
            chunks=self.split_text(test_text,val)
            print("val:",val)
            for i in chunks:
                rectangle_count=rectangle_count+1;
          #  print("rectangle_count :",rectangle_count)

            rectangle_measure=labelSize[0][0]/1140
           # print("rectangle_measure_for text sizw:",rectangle_measure)
            if(len(test_text)<30 and val==1):         
                ori=cv2.rectangle(overlay, (270, 30), (960,100),color_transparent , -1)
                image_new = cv2.addWeighted(overlay, alpha, image, 1 - alpha, 0)
                ori=cv2.rectangle(image_new, (263, 23), (967,107), (255, 255, 255, .4), 10) 
            elif(len(test_text)>=30 and len(test_text)<50 and val==1):
                ori=cv2.rectangle(overlay, (170, 30), (1080,100),color_transparent , -1)
                image_new = cv2.addWeighted(overlay, alpha, image, 1 - alpha, 0)
                ori=cv2.rectangle(image_new, (163, 23), (1087,107), (255, 255, 255, .4), 10)
            elif(len(test_text)<40 and val==2):
                ori=cv2.rectangle(overlay, (170, 30), (1080,100),color_transparent , -1)
                image_new = cv2.addWeighted(overlay, alpha, image, 1 - alpha, 0)
                ori=cv2.rectangle(image_new, (163, 23), (1087,107), (255, 255, 255, .4), 10)
            elif(len(test_text)>=40 and len(test_text)<51 and val==2):
                print("hit rec row 4 ")
                ori=cv2.rectangle(overlay, (50, 30), (1240,100),color_transparent , -1)
                image_new = cv2.addWeighted(overlay, alpha, image, 1 - alpha, 0)
                ori=cv2.rectangle(image_new, (43, 23), (1247,107), (255, 255, 255, .4), 10)
            elif (len(test_text)>=50 and len(test_text)<=74 and val==1):
                ori=cv2.rectangle(overlay, (50, 30), (1240,100),color_transparent , -1)
                image_new = cv2.addWeighted(overlay, alpha, image, 1 - alpha, 0)
                ori=cv2.rectangle(image_new, (43, 23), (1247,107), (255, 255, 255, .4), 10)                  
            elif (rectangle_count==1):
                ori=cv2.rectangle(overlay, (50, 30), (1240,100),color_transparent , -1)
                image_new = cv2.addWeighted(overlay, alpha, image, 1 - alpha, 0)
                ori=cv2.rectangle(image_new, (43, 23), (1247,107), (255, 255, 255, .4), 10)            
            elif (rectangle_count==2):
                ori=cv2.rectangle(overlay, (50, 30), (1240,150),color_transparent , -1)
                image_new = cv2.addWeighted(overlay, alpha, image, 1 - alpha, 0)
                ori=cv2.rectangle(image_new, (43, 23), (1247,157), (255, 255, 255, .4), 10)  
            elif (rectangle_count==3):
                ori=cv2.rectangle(overlay, (50, 30), (1240,200),color_transparent , -1)
                image_new = cv2.addWeighted(overlay, alpha, image, 1 - alpha, 0)
                ori=cv2.rectangle(image_new, (43, 23), (1247,207), (255, 255, 255, .4), 10)  
            elif (rectangle_count==4):
                ori=cv2.rectangle(overlay, (50, 30), (1240,250),color_transparent , -1)
                image_new = cv2.addWeighted(overlay, alpha, image, 1 - alpha, 0)
                ori=cv2.rectangle(image_new, (43, 23), (1247,257), (255, 255, 255, .4), 10)
            elif (rectangle_count==5):
                ori=cv2.rectangle(overlay, (50, 30), (1240,300),color_transparent , -1)
                image_new = cv2.addWeighted(overlay, alpha, image, 1 - alpha, 0)
                ori=cv2.rectangle(image_new, (43, 23), (1247,307), (255, 255, 255, .4), 10)  
            elif (rectangle_count==6):
                ori=cv2.rectangle(overlay, (50, 30), (1240,350),color_transparent , -1)
                image_new = cv2.addWeighted(overlay, alpha, image, 1 - alpha, 0)
                ori=cv2.rectangle(image_new, (43, 23), (1247,357), (255, 255, 255, .4), 10)  
            elif (rectangle_count==7):
                ori=cv2.rectangle(overlay, (50, 30), (1240,400),color_transparent , -1)
                image_new = cv2.addWeighted(overlay, alpha, image, 1 - alpha, 0)
                ori=cv2.rectangle(image_new, (43, 23), (1247,407), (255, 255, 255, .4), 10)  
            elif (rectangle_count==8):
                ori=cv2.rectangle(overlay, (50, 30), (1240,450),color_transparent , -1)
                image_new = cv2.addWeighted(overlay, alpha, image, 1 - alpha, 0)
                ori=cv2.rectangle(image_new, (43, 23), (1247,457), (255, 255, 255, .4), 10)  
            elif (rectangle_count==9):
                ori=cv2.rectangle(overlay, (50, 30), (1240,500),color_transparent , -1)
                image_new = cv2.addWeighted(overlay, alpha, image, 1 - alpha, 0)
                ori=cv2.rectangle(image_new, (43, 23), (1247,507), (255, 255, 255, .4), 10)  
            elif (rectangle_count==10):
                ori=cv2.rectangle(overlay, (50, 30), (1240,550),color_transparent , -1)
                image_new = cv2.addWeighted(overlay, alpha, image, 1 - alpha, 0)
                ori=cv2.rectangle(image_new, (43, 23), (1247,557), (255, 255, 255, .4), 10)  
            elif (rectangle_count==11):
                ori=cv2.rectangle(overlay, (50, 30), (1240,600),color_transparent , -1)
                image_new = cv2.addWeighted(overlay, alpha, image, 1 - alpha, 0)
                ori=cv2.rectangle(image_new, (43, 23), (1247,607), (255, 255, 255, .4), 10)  
            elif (rectangle_count==12):
                ori=cv2.rectangle(overlay, (50, 30), (1240,650),color_transparent , -1)
                image_new = cv2.addWeighted(overlay, alpha, image, 1 - alpha, 0)
                ori=cv2.rectangle(image_new, (43, 23), (1247,657), (255, 255, 255, .4), 10)  
            elif (rectangle_count==13):
                ori=cv2.rectangle(overlay, (50, 30), (1240,700),color_transparent , -1)
                image_new = cv2.addWeighted(overlay, alpha, image, 1 - alpha, 0)
                ori=cv2.rectangle(image_new, (43, 23), (1247,707), (255, 255, 255, .4), 10)  

            print("Length of string",len(test_text))
            length_of_input_string=len(test_text)

             ## font dec;aration here :
            img_pil = Image.fromarray(image_new)
            draw = ImageDraw.Draw(img_pil)  
  
                    # specified font size
            fontpath = r"D:\projects_freelance\Video_creator\fINAL\Cabin-Bold.ttf"
            font = ImageFont.truetype(fontpath, 35)
            #font = ImageFont.truetype("Cabin-Bold.ttf",35)
            #check Capital string here


            if(length_of_input_string>74 and val==1):
                print("value:",val)
                chunks=self.split_text(test_text,val)
                print(chunks);
                count=50;
                for i in chunks:
                    print("chunk:",i)
                    print("Length of Chunk is :",len(i))
 
                    print("count:",count)
                    # drawing text size 
                    draw.text((60, count), i, font = font, align ="left" ,fill=color)
                    img = np.array(img_pil)
                    count=count+50;
                    print("count after :",count)
            elif(length_of_input_string<30 and val==1):
                print("hit this !!! ")
                draw.text((400, 50), test_text, font = font, align ="left",fill=color)  
                img = np.array(img_pil)
            elif(length_of_input_string>=30 and length_of_input_string<50 and val==1):
                # drawing text size
                print("hit this ")
                draw.text((200, 50), test_text, font = font, align ="left",fill=color)  
                img = np.array(img_pil)
            elif( length_of_input_string>=50  and length_of_input_string<=74 and val == 1 ):   
                # drawing text size
                print("lesss than 75 hit _________________________________________________________________!!!!")
                draw.text((60, 50), test_text, font = font, align ="left",fill=color)
                img = np.array(img_pil)
            elif(length_of_input_string>51 and val == 2):
                chunks=self.split_text(test_text,val)
                print(chunks);
                print("hitr me")
                count=50;
                for i in chunks:
                    print("chunk:",i)
                    print("Length of Chunk is :",len(i))
 
                    print("count:",count)
                    # drawing text size 
                    draw.text((60, count), i, font = font, align ="left",fill=color)
                    img = np.array(img_pil)
                    count=count+50;
                    print("count after :",count)
            elif(length_of_input_string>=40 and   length_of_input_string<51 and val == 2):
                print("hit row 4")
                draw.text((60, 50), test_text, font = font, align ="left",fill=color)
                img = np.array(img_pil)
            elif(length_of_input_string<40 and val==2):
                draw.text((200, 50), test_text, font = font, align ="left",fill=color)
                img = np.array(img_pil)               


            img_folder_path=self.temp_path
            output_img=self.Input_dest_path;
            unique_no=datetime.now().timestamp()
            unique_No=str(unique_no)
            d=chr(92)
            out=img_folder_path+d+"img"+unique_No+'.jpg'
            print("des_path:",out)
            
            cv2.imwrite(out,img);
            output_image_path=out
            return output_image_path;
            self.Logger("Images creadted Successfully!!!")
        except Exception as error:
            msg=str(error)
            var = traceback.format_exc()
            self.Logger(var)
            self.setError.emit(msg)
   
        
    def Logger(self,message,level=0):
        #print("Im in Logger")
        d=chr(92)
        file_name=self.Input_dest_path+d+"logs.txt"
        log_file_path=os.path.join(self.Input_dest_path,file_name)
        if os.path.isfile(log_file_path):
            try:
                with open(file_name,'a') as f:
                 ast='**'
                 spaces=ast*level
                 log=spaces+message
                 f.write(str(datetime.now())+' '+str(log)+'\n')
            except Exception as e:
                var = traceback.format_exc()
                self.Logger(var)
            else:
                pass
        else:
            try:
                with open(log_file_path,'w') as f:
                    f.write('Logs for process started on'+str(datetime.now()))
            except:
                var = traceback.format_exc()
                self.Logger(var)
        

# Create Image Functiuons is here
    def path_generator(self,i):
        try: 
            count=i
            img_folder_path=self.temp_path
            output_img=self.Input_dest_path;
            unique_no=datetime.now().timestamp()
            unique_No=str(unique_no)
            d=chr(92)
            count=str(count)
            out=img_folder_path+d+"Row"+count+"video"+".mp4" #temporarry path 
            #print("video_des_path:",out)
            out2=img_folder_path+d+"Rowvideo_aud"+unique_No+".mp4"  #destination path
            out3=output_img+d+"Row"+count+"video"+unique_No+".mp4"  #destination path
            self.Logger("path generator Runs Successfully ")
            return out,out2,out3
        
        except Exception as error:
            var = traceback.format_exc()
            self.Logger(var)
            msg="Input Folder does not exist in yout system"
            self.setError.emit(msg)


    def create_question(self,img_path,i):
        try:

            self.intro_sel_path=random.choice(self.Intro_arr)
            self.outro_sel_path=random.choice(self.Outro_arr)
            print("intro",self.intro_sel_path)
            print("outro",self.outro_sel_path)
            

     
            count=i 
            result=self.path_generator(count)
            out=result[0]
            out2=result[1]
            out3=result[2]
            count_dest=0;
            #set duration
            duration=int(self.Input_time_Perslide)
            
            # Add a  fource
            fourcc = cv2.VideoWriter_fourcc(*'mp4v')
            
            width=1280
            height=720
            video = cv2.VideoWriter(out, fourcc, 1, (width, height)) 
            
            #read Intro image first
            if(self.intro_sel_path.endswith(".jpg") or self.intro_sel_path.endswith(".jpeg") or self.intro_sel_path.endswith(".jfif") or self.intro_sel_path.endswith(".png")):
                image_intro=cv2.imread(self.intro_sel_path)
                image_intro = cv2.resize(image_intro,(1280,720))
                count_dest=count_dest+1;
                for i in range(duration):
                    video.write(image_intro)
            
            #read Questions in video
            for i in img_path:
                for x in range (duration):
                    video.write(cv2.imread(i))
              
            #read outro image
         #   print("outro path :",self.Input_outro_path)
            if(self.outro_sel_path.endswith(".jpg") or self.outro_sel_path.endswith(".jpeg") or self.outro_sel_path.endswith(".jfif") or self.outro_sel_path.endswith(".png")):
                image_outro=cv2.imread(self.outro_sel_path)
                image_outro= cv2.resize(image_outro,(1280,720))
                count_dest=count_dest+1;
                for i in range(duration):
                    video.write(image_outro)
            
            cv2.destroyAllWindows()     
            video.release()
            time.sleep(2)

            self.Logger("Question Videos Sucess!!!!!!!")
            self.Add_Audio(out,out2,out3,count_dest)
            self.concat_video(out,out2,out3)
        except Exception as error:
            var = traceback.format_exc()
            self.Logger(var)
            msg=str(error)
            self.setError.emit(msg)
                   

    def concat_video(self,out,out2,out3):
        try:
            if(self.intro_sel_path.endswith(".mp4") ) and(self.outro_sel_path.endswith(".mp4") ):
                clip1 = VideoFileClip(self.intro_sel_path,audio=True)
                clip3=VideoFileClip(self.outro_sel_path,audio=True)
                clip2=VideoFileClip(out2,audio=True)
                final = concatenate_videoclips([clip1,clip2,clip3],method="compose")
                final.write_videofile(out3)
                clip1.close()
                clip2.close()
                clip3.close()
            if(self.intro_sel_path.endswith(".jpg") or self.intro_sel_path.endswith(".jpeg") or self.intro_sel_path.endswith(".jfif") or self.intro_sel_path.endswith(".png")) and(self.outro_sel_path.endswith(".mp4")):
                clip3 = VideoFileClip(self.outro_sel_path,audio=True)
                clip2=VideoFileClip(out2,audio=True)
                final = concatenate_videoclips([clip2,clip3],method="compose")
                final.write_videofile(out3)
                clip3.reader.close()
                #clip3.audio.reader.close_proc()
                clip2.reader.close()
                clip2.audio.reader.close_proc()
            if(self.intro_sel_path.endswith(".mp4")  ) and(self.outro_sel_path.endswith(".jpg") or self.outro_sel_path.endswith(".png")  or self.outro_sel_path.endswith(".jpeg") or self.outro_sel_path.endswith(".jfif")):
                clip1 = VideoFileClip(self.intro_sel_path,audio=True)
                clip2=VideoFileClip(out2,audio=True)
                final = concatenate_videoclips([clip1,clip2],method="compose")
                print("out2 path:",out2)
                final.write_videofile(out3)
                clip1.close()
                clip2.close()
            self.Logger("Conactenation of Vidoes Success!!")
        except Exception as error:
            var = traceback.format_exc()
            self.Logger(var)           
            msg=str(error)
            self.setError.emit(msg)           
    
    def create_video(self,img_path,i):
        try:
            Image_path=img_path
            count=i
            self.create_question(img_path,i)
            self.Logger("Create video Function Run Successfully !!!!!!!!")
        except Exception as error:
            var = traceback.format_exc()
            self.Logger(var)
            msg=str(error)
            self.setError.emit(msg)
    def Trim_Audio(self,dur):
        try:
            output_img=self.Input_dest_path;
            unique_no=datetime.now().timestamp()
            unique_No=str(unique_no)
            d=chr(92)
            temp_audio=output_img+d+"Temp"+d+"audrow"+unique_No+".mp3" #temporarry path
         #   print("Audio_path:",temp_audio)
            
            file=self.validate_audio()                  #self.Input_audio_path
            if file:
                snd=AudioFileClip(file)
                nsnd=snd.subclip(0,dur)   #'00:00','00:10')
                nsnd.write_audiofile(temp_audio)
                snd.close()
                self.Logger("Trim Audio Success!!!")
                return temp_audio
            else:
                self.Logger("Failed to get a Audio path from Validate Audio Func")
        except Exception as error:
            var = traceback.format_exc()
            self.Logger(var)
            msg=str(error)
            self.setError.emit(msg)
    
    def Add_Audio(self,path1,path2,path3,count_dest):
        try:
            if(count_dest==2):
                path2=path3
          #  print("video_path:",path1)
            vidname= path1  #"C:\\Users\\Lenovo\\Desktop\\Edited_images\\video1.mp4"
            #audname=self.Input_audio_path;
            my_clip = VideoFileClip(vidname)
            dur=my_clip.duration
            s=int(dur)
            trim_aud_path=self.Trim_Audio(s)
            if trim_aud_path:

                audio_background = AudioFileClip(trim_aud_path)
            

            #print("Duration of Video:",my_clip.duration)
            #print("Duration of Audio:",audio_background.duration)

                final_clip = my_clip.set_audio(audio_background)
                #fps=int(self.Aud)
                final_clip.write_videofile(path2,fps=s-1) 
                audio_background.close()
                my_clip.close()
                final_clip.close()
                self.Logger("Add Audio Success")
                
        except Exception as error:
            var = traceback.format_exc()
            self.Logger(var)
            msg="Error occired In Audio try different Audio File "
            self.setError.emit(msg)


    def store_audio(self):
        try:   
            wb = xlrd.open_workbook(self.Excel_filepath) 
            sheet = wb.sheet_by_index(0)
            arr=[]
            rows = sheet.nrows
            columns = sheet.ncols
            self.No_of_Row=columns
            os.chdir(self.Input_audio_path)
            path=self.Input_audio_path   #self.Input_audio_path_first
            limit=0;
            time=int(self.Input_time_Perslide)
            print("time:",time)
            col=columns
            limit=(time*col)+(time*2)
            print("LIMIT:",limit)
            count_audio=0;
            if rows ==0:
                print(self.No_of_Row)
                self.setError.emit("Your Excel File is Empty Enter a Valid Excel File !!!!!! ")
            else:
                audfile=[]
                for root, dirs, files in os.walk(path):
                    for file in files:
                        p=os.path.join(root,file)
                        if p.endswith(".mp3"):
                            audfile.append(os.path.abspath(p))

                size=0;
                for i in audfile:
                    audio_background = AudioFileClip(i)
                    print("Duration of Audio:",audio_background.duration)
                    aud_duration=audio_background.duration
                    print("aud_duration:",aud_duration)
                    size=int(aud_duration)
                    audio_background.close()
                    print("limit:",limit)
                    print("Size:",size)
                    if(limit<size):
                        self.new_audfile.append(i)
                        count_audio=count_audio+1;
                        
            if(count_audio>0):
                print("AudFIles:",self.new_audfile)
                print("--------------------------------------------------------------------------------------")
                self.Logger("Store Valid Audio Files Successfully!!!!")
            else:
                print("else Hit ___________________________________________________________________________________________ else ")
                self.Logger("Error Occured No Audio Files Meet Condition They are smaller ")
                self.setError.emit("Error Occured  Enter a Audio files which has size greater than  "+str(limit)+" seconds")
                self.on_stopprocess(True)                
                
        except Exception as error:
            msg=str(error)
            var = traceback.format_exc()
            self.Logger(var)
            self.setError.emit(msg)

    def store_intro_outro_path(self):
        try:
            print("func intro outro!!!!!")
            os.chdir(self.Input_intro_path)
            path1=self.Input_intro_path   #self.Input_audio_path_first
            arr=[]
            for root, dirs, files in os.walk(path1):
                for file in files:
                    p=os.path.join(root,file)
                    if p.endswith(".jpg") or p.endswith(".jpeg") or p.endswith(".jfif") or p.endswith(".png") or p.endswith(".mp4"):
                        self.Intro_arr.append(os.path.abspath(p))

            #chdir to path2
            os.chdir(self.Input_outro_path)
            path2=self.Input_outro_path
            for root, dirs, files in os.walk(path2):
                for file in files:
                    p=os.path.join(root,file)
                    if p.endswith(".jpg") or p.endswith(".jpeg") or p.endswith(".jfif") or p.endswith(".png") or p.endswith(".mp4"):
                        self.Outro_arr.append(os.path.abspath(p))
            #print("Input Arr:",self.Intro_arr)
            #print("Outro Arr ",self.Outro_arr)
            self.Logger("Intro Outro Files are fetch From Input Folder are Done Sucesssfully!!!!!")
        except Exception as error:
            msg=str(error)
            var = traceback.format_exc()
            self.Logger(var)
            self.setError.emit(msg)

    def validate_audio(self):
        try:
            size=0;

            audpath=random.choice(self.new_audfile)
            limit=self.Limit_aud

            audio_background = AudioFileClip(audpath)
            print("Duration of Audio:",audio_background.duration)
            aud_duration=audio_background.duration
            print("aud_duration:",aud_duration)
            size=int(aud_duration)
            audio_background.close()
            if limit>size:
                while limit>size:
                    audpath=random.choice(self.new_audfile)
                    audio_background = AudioFileClip(audpath)
                    print("Duration of Audio:",audio_background.duration)
                    aud_duration=audio_background.duration
                    print("aud_duration:",aud_duration)
                    size=int(aud_duration)
                    audio_background.close()
                    print("audpath:",audpath)
            
            print("final path:",audpath)
            if size>limit:
                self.Logger("Audio Selection from Given Directory Sucessfully")
                return audpath;
            else:                
                self.Logger("Error Occured No Audio meets Condition Size of Audio is too much less")
                self.setError.emit("Error Occured  Enter a Audio files which has size greater than  "+str(limit)+" seconds")
                self.stop_process=True
        except Exception as error:
            msg=str(error)
            var = traceback.format_exc()
            self.Logger(var)
            self.setError.emit(msg)

    def split_text(self,t,val):
        try:
            if val==1:
                no=74
            if val==2:
                no=53
            t=t;
            Array_of_sentence=[]
            S1=""
            S2=""
            S3=""
            S4=""
            S5=""
            S6=""
            S7=""
            S8=""
            S9=""
            S10=""
            S11=""
            S12=""
            L=0;
            count=0;
            question=t.split()
            for q in question:
                L=len(q)+L+1
                if(L>no):
                    L=0;
                    L=len(q)+1;
                    count=count+1;
                if(count==0 and L<=no):
                    S1=S1+q+" ";
                if(count==1 and L<=no):
                    S2=S2+q+" ";
                if(count==2 and L<=no):
                    S3=S3+q+" ";
                if(count==3 and L<=no):
                    S4=S4+q+" ";
                if(count==4 and L<=no):
                    S5=S5+q+" ";
                if(count==5 and L<=no):
                    S6=S6+q+" ";
                if(count==6 and L<=no):
                    S7=S7+q+" ";
                if(count==7 and L<=no):
                    S8=S8+q+" ";
                if(count==8 and L<=no):
                    S9=S9+q+" ";
                if(count==9 and L<=no):
                    S10=S10+q+" ";
                if(count==10 and L<=no):
                    S11=S11+q+" ";
                if(count==11 and L<=no):
                    S12=S12+q+" ";             
            if(len(S1)>0):
                Array_of_sentence.append(S1)
            if(len(S2)>0):
                Array_of_sentence.append(S2)
            if(len(S3)>0):
                Array_of_sentence.append(S3)
            if(len(S4)>0):
                Array_of_sentence.append(S4)
            if(len(S5)>0):
                Array_of_sentence.append(S5)
            if(len(S6)>0):
                Array_of_sentence.append(S6)
            if(len(S7)>0):
                Array_of_sentence.append(S7)
            if(len(S8)>0):
                Array_of_sentence.append(S8) 
            if(len(S9)>0):
                Array_of_sentence.append(S9) 
            if(len(S10)>0):
                Array_of_sentence.append(S10)
            if(len(S11)>0):
                Array_of_sentence.append(S11)
            if(len(S12)>0):
                Array_of_sentence.append(S12)
            self.Logger("Split the text Function Success!!!")
            return Array_of_sentence;            
        except Exception as error:
            var = traceback.format_exc()
            self.Logger(var)
            msg=str(error)
            self.setError.emit(msg)               


class MyWindow(QMainWindow):
    sig = pyqtSignal(bool)
    Excel_filepath= ""  #GLobal class Variavbles
    Image_folder_path=""
    Input_intro_path=""
    Input_outro_path=""
    Input_audio_path=""
    Input_text_color=""
    Input_back_color=""
    Input_time_Perslide=0
    Input_dest_path=""
    temp_path=""
    filepath=""
    No_of_Row=0
    Error_audio=False
    Aud=0;
    cancel_msg=False
    log_file_path=""
    count_error=False
    Error_Audio_Folder=False

    def __init__(self):
        super(MyWindow,self).__init__()
        Excel_filepath=self.Excel_filepath    #GLobal class Variavbles
        Image_folder_path=self.Image_folder_path
        Input_intro_path=self.Input_intro_path
        Input_outro_path=self.Input_outro_path
        Input_audio_path=self.Input_audio_path
        Input_text_color=self.Input_text_color
        Input_back_color=self.Input_back_color
        Input_time_Perslide=self.Input_time_Perslide
        Input_dest_path=self.Input_dest_path
        temp_path=self.temp_path
        No_of_Row=self.No_of_Row
        Error_audio=self.Error_audio
        Aud=self.Aud
        cancel_msg=self.cancel_msg
        count_error=self.count_error
        Error_Audio_Folder=self.Error_Audio_Folder
        
        self.initUI()
        self.center()
        self.setWindowIcon(QIcon('icons8-commodore-amiga-480.png')) 
        filepath=self.filepath

        self.Input_text_color=(0,0,0)
        self.Input_back_color=(0,0,0)
        #self.Error_Audio_Folder=False

            

    def button_clicked(self):
        self.label.setText("you pressed the button")
        self.update()

    def initUI(self):
        sigstop = pyqtSignal(int)

        self.setWindowFlags(Qt.WindowCloseButtonHint | Qt.WindowMinimizeButtonHint | Qt.CustomizeWindowHint)

        self.threadpool = QThreadPool()    # delette it
        self.setGeometry(200, 200, 800, 580)
        #self.statusBar().showMessage('Ready')
        self.setWindowTitle('Video Creator')
        


        
        #color
        col = QColor(0,0,0) 
               
        #print(self.filepath)

        #self.label = QtWidgets.QLabel(self)
        #self.label.setText("we r we will")
        #self.label.move(300,50)
        
        #line Edits for Every Button
        self.le = QLineEdit(self)
        self.le.move(200, 50)
        self.le.setFixedWidth(520)
        
        
        self.sel_folder=QLineEdit(self)
        self.sel_folder.move(200,100)
        self.sel_folder.setFixedWidth(520)
        
        self.sel_intro=QLineEdit(self)
        self.sel_intro.move(200,150)
        self.sel_intro.setFixedWidth(520)
        
        self.sel_outro=QLineEdit(self)
        self.sel_outro.move(200,200)
        self.sel_outro.setFixedWidth(520)  
        
        self.sel_audio=QLineEdit(self)
        self.sel_audio.move(200,350)
        self.sel_audio.setFixedWidth(520)
        
        self.sel_dest=QLineEdit(self)
        self.sel_dest.move(200,300)
        self.sel_dest.setFixedWidth(520)
        
        self.sel_time=QLineEdit(self)  # timw for one question 
        self.sel_time.move(670,250)
        self.sel_time.setFixedWidth(50)   
        #self.sel_time.setValidator(QIntValidator(1, 20))
        regex=QtCore.QRegExp("/^[1-9]$|^[1-9]$|^1[0-9]$|^20$/")
        validator=QtGui.QRegExpValidator(regex, self.sel_time)
        self.sel_time.setValidator(validator)
        #self.sel_time.returnPressed.connect(self.Validate_audio)
        try:
            self.Input_time_Perslide=self.sel_time.text()
        except Exception as e:
            self.sel_msg.setText("Error Occured Enter Correct  Time")
        #self.Input_time_Perslide=self.sel_time.text()
        

        self.sel_msg=QLineEdit(self)
        self.sel_msg.move(200,490)
        self.sel_msg.setFixedWidth(520)
        
        self.btn_msg = QtWidgets.QPushButton(self)
        self.btn_msg.setEnabled(False)
        self.btn_msg.setText("Messages")
        self.btn_msg.clicked.connect(self.openExcelFileNameDialog)
        self.btn_msg.move(50,490)
        self.btn_msg.setStyleSheet('QPushButton{  background-color: #FF0000; border-radius: 6px;color:white; min-width: 100px;}')
        
        
        
        self.setStyleSheet("""
QProgressBar:horizontal {
    border: 1px solid #3A3939;
    text-align: center;
    padding: 1px;
    background: #201F1F;
}
QProgressBar::chunk:horizontal {
    background-color: qlineargradient(spread:reflect, x1:1, y1:0.545, x2:1, y2:0, stop:0 rgba(28, 66, 111, 255), stop:1 rgba(37, 87, 146, 255));
}

QToolTip
{
    border: 1px solid #3A3939;
    background-color: rgb(90, 102, 117);;
    color: white;
    padding: 1px;
    opacity: 200;
}

QWidget
{
    color: silver;
    background-color: #302F2F;
    selection-background-color:#3d8ec9;
    selection-color: black;
    background-clip: border;
    border-image: none;
    outline: 0;
}

QWidget:item:hover
{
    background-color: #78879b;
    color: black;
}

QWidget:item:selected
{
    background-color: #3d8ec9;
}

QCheckBox
{
    spacing: 5px;
    outline: none;
    color: #bbb;
    margin-bottom: 2px;
}

QCheckBox:disabled
{
    color: #777777;
}
QCheckBox::indicator,
QGroupBox::indicator
{
    width: 18px;
    height: 18px;
}
QGroupBox::indicator
{
    margin-left: 2px;
}

QCheckBox::indicator:unchecked,
QCheckBox::indicator:unchecked:hover,
QGroupBox::indicator:unchecked,
QGroupBox::indicator:unchecked:hover
{
    image: url(:/dark_blue/img/checkbox_unchecked.png);
}

QCheckBox::indicator:unchecked:focus,
QCheckBox::indicator:unchecked:pressed,
QGroupBox::indicator:unchecked:focus,
QGroupBox::indicator:unchecked:pressed
{
  border: none;
    image: url(:/dark_blue/img/checkbox_unchecked_focus.png);
}

QCheckBox::indicator:checked,
QCheckBox::indicator:checked:hover,
QGroupBox::indicator:checked,
QGroupBox::indicator:checked:hover
{
    image: url(:/dark_blue/img/checkbox_checked.png);
}

QCheckBox::indicator:checked:focus,
QCheckBox::indicator:checked:pressed,
QGroupBox::indicator:checked:focus,
QGroupBox::indicator:checked:pressed
{
  border: none;
    image: url(:/dark_blue/img/checkbox_checked_focus.png);
}

QCheckBox::indicator:indeterminate,
QCheckBox::indicator:indeterminate:hover,
QCheckBox::indicator:indeterminate:pressed
QGroupBox::indicator:indeterminate,
QGroupBox::indicator:indeterminate:hover,
QGroupBox::indicator:indeterminate:pressed
{
    image: url(:/dark_blue/img/checkbox_indeterminate.png);
}

QCheckBox::indicator:indeterminate:focus,
QGroupBox::indicator:indeterminate:focus
{
    image: url(:/dark_blue/img/checkbox_indeterminate_focus.png);
}

QCheckBox::indicator:checked:disabled,
QGroupBox::indicator:checked:disabled
{
    image: url(:/dark_blue/img/checkbox_checked_disabled.png);
}

QCheckBox::indicator:unchecked:disabled,
QGroupBox::indicator:unchecked:disabled
{
    image: url(:/dark_blue/img/checkbox_unchecked_disabled.png);
}

QRadioButton
{
    spacing: 5px;
    outline: none;
    color: #bbb;
    margin-bottom: 2px;
}

QRadioButton:disabled
{
    color: #777777;
}
QRadioButton::indicator
{
    width: 21px;
    height: 21px;
}

QRadioButton::indicator:unchecked,
QRadioButton::indicator:unchecked:hover
{
    image: url(:/dark_blue/img/radio_unchecked.png);
}

QRadioButton::indicator:unchecked:focus,
QRadioButton::indicator:unchecked:pressed
{
  border: none;
  outline: none;
    image: url(:/dark_blue/img/radio_unchecked_focus.png);
}

QRadioButton::indicator:checked,
QRadioButton::indicator:checked:hover
{
  border: none;
  outline: none;
    image: url(:/dark_blue/img/radio_checked.png);
}

QRadioButton::indicator:checked:focus,
QRadioButton::indicato::menu-arrowr:checked:pressed
{
  border: none;
  outline: none;
    image: url(:/dark_blue/img/radio_checked_focus.png);
}

QRadioButton::indicator:indeterminate,
QRadioButton::indicator:indeterminate:hover,
QRadioButton::indicator:indeterminate:pressed
{
        image: url(:/dark_blue/img/radio_indeterminate.png);
}

QRadioButton::indicator:checked:disabled
{
  outline: none;
  image: url(:/dark_blue/img/radio_checked_disabled.png);
}

QRadioButton::indicator:unchecked:disabled
{
    image: url(:/dark_blue/img/radio_unchecked_disabled.png);
}


QMenuBar
{
    background-color: #302F2F;
    color: silver;
}

QMenuBar::item
{
    background: transparent;
}

QMenuBar::item:selected
{
    background: transparent;
    border: 1px solid #3A3939;
}

QMenuBar::item:pressed
{
    border: 1px solid #3A3939;
    background-color: #3d8ec9;
    color: black;
    margin-bottom:-1px;
    padding-bottom:1px;
}

QMenu
{
    border: 1px solid #3A3939;
    color: silver;
    margin: 1px;
}

QMenu::icon
{
    margin: 1px;
}

QMenu::item
{
    padding: 2px 2px 2px 25px;
    margin-left: 5px;
    border: 1px solid transparent; /* reserve space for selection border */
}

QMenu::item:selected
{
    color: black;
}

QMenu::separator {
    height: 2px;
    background: lightblue;
    margin-left: 10px;
    margin-right: 5px;
}

QMenu::indicator {
    width: 16px;
    height: 16px;
}

/* non-exclusive indicator = check box style indicator
   (see QActionGroup::setExclusive) */
QMenu::indicator:non-exclusive:unchecked {
    image: url(:/dark_blue/img/checkbox_unchecked.png);
}

QMenu::indicator:non-exclusive:unchecked:selected {
    image: url(:/dark_blue/img/checkbox_unchecked_disabled.png);
}

QMenu::indicator:non-exclusive:checked {
    image: url(:/dark_blue/img/checkbox_checked.png);
}

QMenu::indicator:non-exclusive:checked:selected {
    image: url(:/dark_blue/img/checkbox_checked_disabled.png);
}

/* exclusive indicator = radio button style indicator (see QActionGroup::setExclusive) */
QMenu::indicator:exclusive:unchecked {
    image: url(:/dark_blue/img/radio_unchecked.png);
}

QMenu::indicator:exclusive:unchecked:selected {
    image: url(:/dark_blue/img/radio_unchecked_disabled.png);
}

QMenu::indicator:exclusive:checked {
    image: url(:/dark_blue/img/radio_checked.png);
}

QMenu::indicator:exclusive:checked:selected {
    image: url(:/dark_blue/img/radio_checked_disabled.png);
}

QMenu::right-arrow {
    margin: 5px;
    image: url(:/dark_blue/img/right_arrow.png)
}


QWidget:disabled
{
    color: #808080;
    background-color: #302F2F;
}

QAbstractItemView
{
    alternate-background-color: #3A3939;
    color: silver;
    border: 1px solid 3A3939;
    border-radius: 2px;
    padding: 1px;
}

QWidget:focus, QMenuBar:focus
{
    border: 1px solid #78879b;
}

QTabWidget:focus, QCheckBox:focus, QRadioButton:focus, QSlider:focus
{
    border: none;
}

QLineEdit
{
    background-color: #201F1F;
    padding: 2px;
    border-style: solid;
    border: 1px solid #3A3939;
    border-radius: 10px;
}

QGroupBox {
    border:1px solid #3A3939;
    border-radius: 2px;
    margin-top: 20px;
    background-color: #302F2F;
    color: silver;
}

QGroupBox::title {
    subcontrol-origin: margin;
    subcontrol-position: top center;
    padding-left: 10px;
    padding-right: 10px;
    padding-top: 10px;
}

QAbstractScrollArea
{
    border-radius: 2px;
    border: 1px solid #3A3939;
    background-color: transparent;
}

QScrollBar:horizontal
{
    height: 15px;
    margin: 3px 15px 3px 15px;
    border: 1px transparent #2A2929;
    border-radius: 4px;
    background-color: #2A2929;
}

QScrollBar::handle:horizontal
{
    background-color: #605F5F;
    min-width: 5px;
    border-radius: 4px;
}

QScrollBar::add-line:horizontal
{
    margin: 0px 3px 0px 3px;
    border-image: url(:/dark_blue/img/right_arrow_disabled.png);
    width: 10px;
    height: 10px;
    subcontrol-position: right;
    subcontrol-origin: margin;
}

QScrollBar::sub-line:horizontal
{
    margin: 0px 3px 0px 3px;
    border-image: url(:/dark_blue/img/left_arrow_disabled.png);
    height: 10px;
    width: 10px;
    subcontrol-position: left;
    subcontrol-origin: margin;
}

QScrollBar::add-line:horizontal:hover,QScrollBar::add-line:horizontal:on
{
    border-image: url(:/dark_blue/img/right_arrow.png);
    height: 10px;
    width: 10px;
    subcontrol-position: right;
    subcontrol-origin: margin;
}


QScrollBar::sub-line:horizontal:hover, QScrollBar::sub-line:horizontal:on
{
    border-image: url(:/dark_blue/img/left_arrow.png);
    height: 10px;
    width: 10px;
    subcontrol-position: left;
    subcontrol-origin: margin;
}

QScrollBar::up-arrow:horizontal, QScrollBar::down-arrow:horizontal
{
    background: none;
}


QScrollBar::add-page:horizontal, QScrollBar::sub-page:horizontal
{
    background: none;
}

QScrollBar:vertical
{
    background-color: #2A2929;
    width: 15px;
    margin: 15px 3px 15px 3px;
    border: 1px transparent #2A2929;
    border-radius: 4px;
}

QScrollBar::handle:vertical
{
    background-color: #605F5F;
    min-height: 5px;
    border-radius: 4px;
}

QScrollBar::sub-line:vertical
{
    margin: 3px 0px 3px 0px;
    border-image: url(:/dark_blue/img/up_arrow_disabled.png);
    height: 10px;
    width: 10px;
    subcontrol-position: top;
    subcontrol-origin: margin;
}

QScrollBar::add-line:vertical
{
    margin: 3px 0px 3px 0px;
    border-image: url(:/dark_blue/img/down_arrow_disabled.png);
    height: 10px;
    width: 10px;
    subcontrol-position: bottom;
    subcontrol-origin: margin;
}

QScrollBar::sub-line:vertical:hover,QScrollBar::sub-line:vertical:on
{

    border-image: url(:/dark_blue/img/up_arrow.png);
    height: 10px;
    width: 10px;
    subcontrol-position: top;
    subcontrol-origin: margin;
}


QScrollBar::add-line:vertical:hover, QScrollBar::add-line:vertical:on
{
    border-image: url(:/dark_blue/img/down_arrow.png);
    height: 10px;
    width: 10px;
    subcontrol-position: bottom;
    subcontrol-origin: margin;
}

QScrollBar::up-arrow:vertical, QScrollBar::down-arrow:vertical
{
    background: none;
}


QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical
{
    background: none;
}

QTextEdit
{
    background-color: #201F1F;
    color: silver;
    border: 1px solid #3A3939;
}

QPlainTextEdit
{
    background-color: #201F1F;;
    color: silver;
    border-radius: 2px;
    border: 1px solid #3A3939;
}

QHeaderView::section
{
    background-color: #3A3939;
    color: silver;
    padding-left: 4px;
    border: 1px solid #6c6c6c;
}

QSizeGrip {
    image: url(:/dark_blue/img/sizegrip.png);
    width: 12px;
    height: 12px;
}

QMainWindow
{
    background-color: #302F2F;

}

QMainWindow::separator
{
    background-color: #302F2F;
    color: white;
    padding-left: 4px;
    spacing: 2px;
    border: 1px dashed #3A3939;
}

QMainWindow::separator:hover
{

    background-color: #787876;
    color: white;
    padding-left: 4px;
    border: 1px solid #3A3939;
    spacing: 2px;
}


QMenu::separator
{
    height: 1px;
    background-color: #3A3939;
    color: white;
    padding-left: 4px;
    margin-left: 10px;
    margin-right: 5px;
}


QFrame
{
    border-radius: 8px;
    border: 1px solid #444;
    padding:2px;
}

QFrame[frameShape="0"]
{
    border-radius: 8px;
    border: 1px transparent #444;
}

QStackedWidget
{
    background-color: #302F2F;
    border: 1px transparent black;
}

QToolBar {
    border: 1px transparent #393838;
    background: 1px solid #302F2F;
    font-weight: bold;
}

QToolBar::handle:horizontal {
    image: url(:/dark_blue/img/Hmovetoolbar.png);
}
QToolBar::handle:vertical {
    image: url(:/dark_blue/img/Vmovetoolbar.png);
}
QToolBar::separator:horizontal {
    image: url(:/dark_blue/img/Hsepartoolbar.png);
}
QToolBar::separator:vertical {
    image: url(:/dark_blue/img/Vsepartoolbars.png);
}

QPushButton
{
    color: silver;
    background-color: #302F2F;
    border-width: 2px;
    border-color: #4A4949;
    border-style: solid;
    padding-top: 2px;
    padding-bottom: 2px;
    padding-left: 10px;
    padding-right: 10px;
    border-radius: 4px;
    /* outline: none; */
    /* min-width: 40px; */
}

QPushButton:disabled
{
    background-color: #302F2F;
    border-width: 2px;
    border-color: #3A3939;
    border-style: solid;
    padding-top: 2px;
    padding-bottom: 2px;
    padding-left: 10px;
    padding-right: 10px;
    /*border-radius: 2px;*/
    color: #808080;
}

QPushButton:focus {
    background-color: #3d8ec9;
    color: white;
}

QComboBox
{
    selection-background-color: #3d8ec9;
    background-color: #201F1F;
    border-style: solid;
    border: 1px solid #3A3939;
    border-radius: 2px;
    padding: 2px;
    min-width: 75px;
}

QPushButton:checked{
    background-color: #4A4949;
    border-color: #6A6969;
}

QPushButton:hover {
    border: 2px solid #78879b;
    color: silver;
}

QComboBox:hover, QAbstractSpinBox:hover,QLineEdit:hover,QTextEdit:hover,QPlainTextEdit:hover,QAbstractView:hover,QTreeView:hover
{
    border: 1px solid #78879b;
    color: silver;
}

QComboBox:on
{
    background-color: #626873;
    padding-top: 3px;
    padding-left: 4px;
    selection-background-color: #4a4a4a;
}

QComboBox QAbstractItemView
{
    background-color: #201F1F;
    border-radius: 2px;
    border: 1px solid #444;
    selection-background-color: #3d8ec9;
    color: silver;
}

QComboBox::drop-down
{
    subcontrol-origin: padding;
    subcontrol-position: top right;
    width: 15px;

    border-left-width: 0px;
    border-left-color: darkgray;
    border-left-style: solid;
    border-top-right-radius: 3px;
    border-bottom-right-radius: 3px;
}

QComboBox::down-arrow
{
    image: url(:/dark_blue/img/down_arrow_disabled.png);
}

QComboBox::down-arrow:on, QComboBox::down-arrow:hover,
QComboBox::down-arrow:focus
{
    image: url(:/dark_blue/img/down_arrow.png);
}

QPushButton:pressed
{
    background-color: #484846;
}

QAbstractSpinBox {
    padding-top: 2px;
    padding-bottom: 2px;
    border: 1px solid #3A3939;
    background-color: #201F1F;
    color: silver;
    border-radius: 2px;
    min-width: 75px;
}

QAbstractSpinBox:up-button
{
    background-color: transparent;
    subcontrol-origin: border;
    subcontrol-position: top right;
}

QAbstractSpinBox:down-button
{
    background-color: transparent;
    subcontrol-origin: border;
    subcontrol-position: bottom right;
}

QAbstractSpinBox::up-arrow,QAbstractSpinBox::up-arrow:disabled,QAbstractSpinBox::up-arrow:off {
    image: url(:/dark_blue/img/up_arrow_disabled.png);
    width: 10px;
    height: 10px;
}
QAbstractSpinBox::up-arrow:hover
{
    image: url(:/dark_blue/img/up_arrow.png);
}


QAbstractSpinBox::down-arrow,QAbstractSpinBox::down-arrow:disabled,QAbstractSpinBox::down-arrow:off
{
    image: url(:/dark_blue/img/down_arrow_disabled.png);
    width: 10px;
    height: 10px;
}
QAbstractSpinBox::down-arrow:hover
{
    image: url(:/dark_blue/img/down_arrow.png);
}


QLabel
{
    border: 0px solid black;
}

QTabWidget{
    border: 1px transparent black;
}

QTabWidget::pane {
    border: 1px solid #444;
    border-radius: 3px;
    padding: 3px;
}

QTabBar
{
    qproperty-drawBase: 0;
    left: 5px; /* move to the right by 5px */
}

QTabBar:focus
{
    border: 0px transparent black;
}

QTabBar::close-button  {
    image: url(:/dark_blue/img/close.png);
    background: transparent;
}

QTabBar::close-button:hover
{
    image: url(:/dark_blue/img/close-hover.png);
    background: transparent;
}

QTabBar::close-button:pressed {
    image: url(:/dark_blue/img/close-pressed.png);
    background: transparent;
}

/* TOP TABS */
QTabBar::tab:top {
    color: #b1b1b1;
    border: 1px solid #4A4949;
    border-bottom: 1px transparent black;
    background-color: #302F2F;
    padding: 5px;
    border-top-left-radius: 2px;
    border-top-right-radius: 2px;
}

QTabBar::tab:top:!selected
{
    color: #b1b1b1;
    background-color: #201F1F;
    border: 1px transparent #4A4949;
    border-bottom: 1px transparent #4A4949;
    border-top-left-radius: 0px;
    border-top-right-radius: 0px;
}

QTabBar::tab:top:!selected:hover {
    background-color: #48576b;
}

/* BOTTOM TABS */
QTabBar::tab:bottom {
    color: #b1b1b1;
    border: 1px solid #4A4949;
    border-top: 1px transparent black;
    background-color: #302F2F;
    padding: 5px;
    border-bottom-left-radius: 2px;
    border-bottom-right-radius: 2px;
}

QTabBar::tab:bottom:!selected
{
    color: #b1b1b1;
    background-color: #201F1F;
    border: 1px transparent #4A4949;
    border-top: 1px transparent #4A4949;
    border-bottom-left-radius: 0px;
    border-bottom-right-radius: 0px;
}

QTabBar::tab:bottom:!selected:hover {
    background-color: #78879b;
}

/* LEFT TABS */
QTabBar::tab:left {
    color: #b1b1b1;
    border: 1px solid #4A4949;
    border-left: 1px transparent black;
    background-color: #302F2F;
    padding: 5px;
    border-top-right-radius: 2px;
    border-bottom-right-radius: 2px;
}

QTabBar::tab:left:!selected
{
    color: #b1b1b1;
    background-color: #201F1F;
    border: 1px transparent #4A4949;
    border-right: 1px transparent #4A4949;
    border-top-right-radius: 0px;
    border-bottom-right-radius: 0px;
}

QTabBar::tab:left:!selected:hover {
    background-color: #48576b;
}


/* RIGHT TABS */
QTabBar::tab:right {
    color: #b1b1b1;
    border: 1px solid #4A4949;
    border-right: 1px transparent black;
    background-color: #302F2F;
    padding: 5px;
    border-top-left-radius: 2px;
    border-bottom-left-radius: 2px;
}

QTabBar::tab:right:!selected
{
    color: #b1b1b1;
    background-color: #201F1F;
    border: 1px transparent #4A4949;
    border-right: 1px transparent #4A4949;
    border-top-left-radius: 0px;
    border-bottom-left-radius: 0px;
}

QTabBar::tab:right:!selected:hover {
    background-color: #48576b;
}

QTabBar QToolButton::right-arrow:enabled {
     image: url(:/dark_blue/img/right_arrow.png);
 }

 QTabBar QToolButton::left-arrow:enabled {
     image: url(:/dark_blue/img/left_arrow.png);
 }

QTabBar QToolButton::right-arrow:disabled {
     image: url(:/dark_blue/img/right_arrow_disabled.png);
 }

 QTabBar QToolButton::left-arrow:disabled {
     image: url(:/dark_blue/img/left_arrow_disabled.png);
 }


QDockWidget {
    border: 1px solid #403F3F;
    titlebar-close-icon: url(:/dark_blue/img/close.png);
    titlebar-normal-icon: url(:/dark_blue/img/undock.png);
}

QDockWidget::close-button, QDockWidget::float-button {
    border: 1px solid transparent;
    border-radius: 2px;
    background: transparent;
}

QDockWidget::close-button:hover, QDockWidget::float-button:hover {
    background: rgba(255, 255, 255, 10);
}

QDockWidget::close-button:pressed, QDockWidget::float-button:pressed {
    padding: 1px -1px -1px 1px;
    background: rgba(255, 255, 255, 10);
}

QTreeView, QListView, QTextBrowser, AtLineEdit, AtLineEdit::hover {
    border: 1px solid #444;
    background-color: silver;
    border-radius: 3px;
    margin-left: 3px;
    color: black;
}

QTreeView:branch:selected, QTreeView:branch:hover {
    background: url(:/dark_blue/img/transparent.png);
}

QTreeView::branch:has-siblings:!adjoins-item {
    border-image: url(:/dark_blue/img/transparent.png);
}

QTreeView::branch:has-siblings:adjoins-item {
    border-image: url(:/dark_blue/img/transparent.png);
}

QTreeView::branch:!has-children:!has-siblings:adjoins-item {
    border-image: url(:/dark_blue/img/transparent.png);
}

QTreeView::branch:has-children:!has-siblings:closed,
QTreeView::branch:closed:has-children:has-siblings {
    image: url(:/dark_blue/img/branch_closed.png);
}

QTreeView::branch:open:has-children:!has-siblings,
QTreeView::branch:open:has-children:has-siblings  {
    image: url(:/dark_blue/img/branch_open.png);
}

QTreeView::branch:has-children:!has-siblings:closed:hover,
QTreeView::branch:closed:has-children:has-siblings:hover {
    image: url(:/dark_blue/img/branch_closed-on.png);
    }

QTreeView::branch:open:has-children:!has-siblings:hover,
QTreeView::branch:open:has-children:has-siblings:hover  {
    image: url(:/dark_blue/img/branch_open-on.png);
    }

QListView::item:!selected:hover, QListView::item:!selected:hover, QTreeView::item:!selected:hover  {
    background: rgba(0, 0, 0, 0);
    outline: 0;
    color: #FFFFFF
}

QListView::item:selected:hover, QListView::item:selected:hover, QTreeView::item:selected:hover  {
    background: #3d8ec9;
    color: #FFFFFF;
}

QSlider::groove:horizontal {
    border: 1px solid #3A3939;
    height: 8px;
    background: #201F1F;
    margin: 2px 0;
    border-radius: 2px;
}

QSlider::handle:horizontal {
    background: QLinearGradient( x1: 0, y1: 0, x2: 0, y2: 1,
      stop: 0.0 silver, stop: 0.2 #a8a8a8, stop: 1 #727272);
    border: 1px solid #3A3939;
    width: 14px;
    height: 14px;
    margin: -4px 0;
    border-radius: 2px;
}

QSlider::groove:vertical {
    border: 1px solid #3A3939;
    width: 8px;
    background: #201F1F;
    margin: 0 0px;
    border-radius: 2px;
}

QSlider::handle:vertical {
    background: QLinearGradient( x1: 0, y1: 0, x2: 0, y2: 1, stop: 0.0 silver,
    stop: 0.2 #a8a8a8, stop: 1 #727272);
    border: 1px solid #3A3939;
    width: 14px;
    height: 14px;
    margin: 0 -4px;
    border-radius: 2px;
}

QToolButton {
    /*  background-color: transparent; */
    border: 2px transparent #4A4949;
    border-radius: 4px;
    background-color: dimgray;
    margin: 2px;
    padding: 2px;
}

QToolButton[popupMode="1"] { /* only for MenuButtonPopup */
 padding-right: 20px; /* make way for the popup button */
 border: 2px transparent #4A4949;
 border-radius: 4px;
}

QToolButton[popupMode="2"] { /* only for InstantPopup */
 padding-right: 10px; /* make way for the popup button */
 border: 2px transparent #4A4949;
}


QToolButton:hover, QToolButton::menu-button:hover {
    border: 2px solid #78879b;
}

QToolButton:checked, QToolButton:pressed,
    QToolButton::menu-button:pressed {
    background-color: #4A4949;
    border: 2px solid #78879b;
}

/* the subcontrol below is used only in the InstantPopup or DelayedPopup mode */
QToolButton::menu-indicator {
    image: url(:/dark_blue/img/down_arrow.png);
    top: -7px; left: -2px; /* shift it a bit */
}

/* the subcontrols below are used only in the MenuButtonPopup mode */
QToolButton::menu-button {
    border: 1px transparent #4A4949;
    border-top-right-radius: 6px;
    border-bottom-right-radius: 6px;
    /* 16px width + 4px for border = 20px allocated above */
    width: 16px;
    outline: none;
}

QToolButton::menu-arrow {
    image: url(:/dark_blue/img/down_arrow.png);
}

QToolButton::menu-arrow:open {
    top: 1px; left: 1px; /* shift it a bit */
    border: 1px solid #3A3939;
}

QPushButton::menu-indicator  {
    subcontrol-origin: padding;
    subcontrol-position: bottom right;
    left: 4px;
}

QTableView
{
    border: 1px solid #444;
    gridline-color: #6c6c6c;
    background-color: #201F1F;
}


QTableView, QHeaderView
{
    border-radius: 0px;
}

QTableView::item:pressed, QListView::item:pressed, QTreeView::item:pressed  {
    background: #78879b;
    color: #FFFFFF;
}

QTableView::item:selected:active, QTreeView::item:selected:active, QListView::item:selected:active  {
    background: #3d8ec9;
    color: #FFFFFF;
}


QHeaderView
{
    border: 1px transparent;
    border-radius: 2px;
    margin: 0px;
    padding: 0px;
}

QHeaderView::section  {
    background-color: #3A3939;
    color: silver;
    padding: 4px;
    border: 1px solid #6c6c6c;
    border-radius: 0px;
    text-align: center;
}

QHeaderView::section::vertical::first, QHeaderView::section::vertical::only-one
{
    border-top: 1px solid #6c6c6c;
}

QHeaderView::section::vertical
{
    border-top: transparent;
}

QHeaderView::section::horizontal::first, QHeaderView::section::horizontal::only-one
{
    border-left: 1px solid #6c6c6c;
}

QHeaderView::section::horizontal
{
    border-left: transparent;
}


QHeaderView::section:checked
 {
    color: white;
    background-color: #5A5959;
 }

 /* style the sort indicator */
QHeaderView::down-arrow {
    image: url(:/dark_blue/img/down_arrow.png);
}

QHeaderView::up-arrow {
    image: url(:/dark_blue/img/up_arrow.png);
}


QTableCornerButton::section {
    background-color: #3A3939;
    border: 1px solid #3A3939;
    border-radius: 2px;
}

QToolBox  {
    padding: 3px;
    border: 1px transparent black;
}

QToolBox::tab {
    color: #b1b1b1;
    background-color: #302F2F;
    border: 1px solid #4A4949;
    border-bottom: 1px transparent #302F2F;
    border-top-left-radius: 5px;
    border-top-right-radius: 5px;
}

 QToolBox::tab:selected { /* italicize selected tabs */
    font: italic;
    background-color: #302F2F;
    border-color: #3d8ec9;
 }

QStatusBar::item {
    border: 1px solid #3A3939;
    border-radius: 2px;
 }


QSplitter::handle {
    border: 1px dashed #3A3939;
}

QSplitter::handle:hover {
    background-color: #787876;
    border: 1px solid #3A3939;
}

QSplitter::handle:horizontal {
    width: 1px;
}

QSplitter::handle:vertical {
    height: 1px;
}

QListWidget {
    background-color: silver;
    border-radius: 5px;
    margin-left: 5px;
}

QListWidget::item {
    color: black;
}

QMessageBox {
    messagebox-critical-icon	: url(:/dark_blue/img/critical.png);
    messagebox-information-icon	: url(:/dark_blue/img/information.png);
    messagebox-question-icon	: url(:/dark_blue/img/question.png);
    messagebox-warning-icon:    : url(:/dark_blue/img/warning.png);
    min-width:500 px;
    font-size: 20px;
    QMessageBox.Yes{wisth:40;height:20}

    
}

ColorButton::enabled {
    border-radius: 0px;
    border: 1px solid #444444;
}

ColorButton::disabled {
    border-radius: 0px;
    border: 1px solid #AAAAAA;
}

        """)
        
        


        self.progress = QProgressBar(self)
        self.progress.setGeometry(50, 540, 700, 30)
        self.progress.setMaximum(100)

        
        
        
        
        #self.btn_file.setObjectName("first_button");

        self.btn_file = QtWidgets.QPushButton(self)
        self.btn_file.setText("Add File")
        self.btn_file.clicked.connect(self.openExcelFileNameDialog)
        self.btn_file.move(50,50)
        self.btn_file.setStyleSheet('QPushButton{  border-radius: 6px;min-width: 100px;}')
        
        self.btn_img_folder = QtWidgets.QPushButton(self)
        self.btn_img_folder.setText("Select Image Folder")
        self.btn_img_folder.clicked.connect(self.openFolder)
        self.btn_img_folder.move(50,100)
        self.btn_img_folder.setStyleSheet('QPushButton{ border-radius: 6px;min-width: 100px;}')
        
        self.btn_intro = QtWidgets.QPushButton(self)
        self.btn_intro.setText("Add Introduction")
        self.btn_intro.clicked.connect(self.OpenIntroDialog)
        self.btn_intro.move(50,150)
        self.btn_intro.setStyleSheet('QPushButton{ border-radius: 6px;min-width: 100px;}')
        
        self.btn_outro = QtWidgets.QPushButton(self)
        self.btn_outro.setText("Add Outro ")
        self.btn_outro.clicked.connect(self.OpenOutroDialog)
        self.btn_outro.move(50,200)  
        self.btn_outro.setStyleSheet('QPushButton{ border-radius: 6px;min-width: 100px;}')
          
        self.btn_audio = QtWidgets.QPushButton(self)
        self.btn_audio.setText("Add Audio")
        self.btn_audio.clicked.connect(self.openFolder_audio)
        self.btn_audio.move(50,350)
        self.btn_audio.setStyleSheet('QPushButton{ border-radius: 6px;min-width: 100px;}')
        
        self.btn_start = QtWidgets.QPushButton(self)
        self.btn_start.setObjectName('button1')
        self.btn_start.setText("Start⯈")
        self.btn_start.clicked.connect(self.Start)
        #self.btn_start.clicked.connect(self.onButtonClick)
        self.btn_start.move(300,405)#430)
        self.btn_start.setStyleSheet('QPushButton{ min-height: 40px; min-width: 150px;font-size: 20px;color:white;font-family: "Times New Roman";font-weight: bold;border:1px solid white;border-radius: 8px;}')
     
        self.btn_cancel = QtWidgets.QPushButton(self)
        self.btn_cancel.setObjectName('button1')
        self.btn_cancel.setText("Cancel")
        #self.btn_start.clicked.connect(self.OpenAudioDialog)
        self.btn_cancel.move(530,405)#430)
        self.btn_cancel.clicked.connect(self.onstop_work) # new one
        self.btn_cancel.setEnabled(False)
        self.btn_cancel.setStyleSheet('QPushButton{ min-height: 40px; min-width: 150px;font-size: 20px;font-family: "Times New Roman";font-weight: bold;border:1px solid white;border-radius: 8px;}')
        
        self.btn_destination = QtWidgets.QPushButton(self)
        self.btn_destination.setText("Destination Folder")
        self.btn_destination.clicked.connect(self.openFolder_dest)
        self.btn_destination.move(50,300)  
        self.btn_destination.setStyleSheet('QPushButton{ border-radius: 6px;min-width: 100px;}')
          
 
        # Color Add Area For text 
        self.btn_text_color = QtWidgets.QPushButton(self)
        self.btn_text_color.setText("Add Text Color")
        self.btn_text_color.clicked.connect(self.showColorDialog)
        self.btn_text_color.move(50,250)
        self.btn_text_color.setStyleSheet('QPushButton{ border-radius: 6px;min-width: 100px;}')
                
        
        #open color chooser
        self.frm = QFrame(self)
        self.frm.setStyleSheet("QWidget { background-color: %s;border-radius:8px; }" 
            % col.name())
        self.frm.setGeometry(200, 250, 60, 30)
        
        self.frm.setToolTip('Text Color')

        
        #backgound-color 
        self.btn_back_color = QtWidgets.QPushButton(self)
        self.btn_back_color.setText("Backgound Color")
        self.btn_back_color.clicked.connect(self.showColorDialog2)
        self.btn_back_color.move(280,250)
        self.btn_back_color.setStyleSheet('QPushButton{ border-radius: 6px;min-width: 100px;}')
                
        
        #open color chooser
        self.frm1 = QFrame(self)
        self.frm1.setStyleSheet("QWidget { background-color: %s;border-radius:8px; }" 
            % col.name())
        self.frm1.setGeometry(430, 250, 60, 30)
        self.frm1.setToolTip('Text Background Color')
        
        self.btn_time = QtWidgets.QPushButton(self)
        self.btn_time.setText("Duration of Questions")
        #self.btn_time.clicked.connect(self.showColorDialog2)
        self.btn_time.move(520,250)
        self.btn_time.setStyleSheet('QPushButton{ border-radius: 6px;min-width: 110px;}')
        self.btn_time.setEnabled(False)
        
        

    def center(self):
        qr = self.frameGeometry()
        cp = QDesktopWidget().availableGeometry().center()
        qr.moveCenter(cp)
        self.move(qr.topLeft())
        
        
    def closeEvent(self, event):
        reply = QMessageBox.question(self, 'Message',
            "Are you sure to quit?", QMessageBox.Yes | 
            QMessageBox.No, QMessageBox.No)

        

        if reply == QMessageBox.Yes:
            event.accept()
        else:
            event.ignore()


    def openFolder_audio(self):
        try:
            file = str(QFileDialog.getExistingDirectory(self, "Select Audio Folder"))
            self.Input_audio_path=file
            self.sel_audio.setText(str(file))
        except Exception as e:
            self.sel_msg.setText("Error Occured No Folder/Directory Exist !!!!")              

    def openFolder_dest(self):
        try:
            file = str(QFileDialog.getExistingDirectory(self, "Select Destination Folder"))
            self.Input_dest_path=file
            self.sel_dest.setText(str(file))
        except Exception as e:
            self.sel_msg.setText("Error Occured No Folder/Directory Exist !!!!")        
    def openFolder(self):
        try:
            file = str(QFileDialog.getExistingDirectory(self, "Select Image Folder"))
            self.Image_folder_path=file;
            self.sel_folder.setText(str(file))
        except Exception as e:
            self.sel_msg.setText("Error Occured No Folder/Directory Exist !!!!")
    def OpenIntroDialog(self):
        try:
            file = str(QFileDialog.getExistingDirectory(self, "Select Intro Folder"))
            self.Input_intro_path=file;
            self.sel_intro.setText(str(file))
        except Exception as e:
            self.sel_msg.setText("Error Occured No Folder/Directory Exist !!!!")
                
    def OpenOutroDialog(self):
        try:
            file = str(QFileDialog.getExistingDirectory(self, "Select Outro Folder"))
            self.Input_outro_path=file;
            self.sel_outro.setText(str(file))
        except Exception as e:
            self.sel_msg.setText("Error Occured No Folder/Directory Exist !!!!")
                
                        
    def openExcelFileNameDialog(self):
        try:
            options = QFileDialog.Options()
            fileName, _ = QFileDialog.getOpenFileName(self,"Select Excel File", "","Excel Files (*.xlsx)", options=options)
            self.Excel_filepath =fileName
            if fileName:
                self.filepath=fileName
                self.le.setText(str(fileName))
        except Exception as e:
            self.sel_msg.setText("Error Occured No File  Exist !!!!")
            
    def showColorDialog(self):
        try: 
            col = QColorDialog.getColor()
            if col.isValid():
                self.frm.setStyleSheet("QWidget { background-color: %s }"
                    % col.name())
            h=ImageColor.getrgb(col.name())  
            self.Input_text_color=h
            print(h)
        except Exception as e:
            self.sel_msg.setText("Error Occured Select Color !!!!")
        
    def showColorDialog2(self):
        try:
            col = QColorDialog.getColor()
            if col.isValid():
                self.frm1.setStyleSheet("QWidget { background-color: %s }"
                    % col.name())
            h=ImageColor.getrgb(col.name())   #print(h)
            self.Input_back_color=h
        except Exception as e:
            self.sel_msg.setText("Error Occured Select Color!!!!")
        
    def Start(self):
        try:
            self.sel_msg.setText(" ")
            self.progress.setValue(0)
            print("END_Start")
            #self.Validate_audio()
            self.validate_input()
        except Exception as e:
            val=str(e);
            msg="Error ocuured "+val+" "
            self.sel_msg.setText(msg)

    def RepresentsInt(self,s):
        try: 
            int(s)
            return True
        except ValueError:
            return False
     
##    def Validate_audio(self):
##        try:
##            self.Error_Audio_Folder=False
##            print("error_audio---",self.Error_Audio_Folder)
##            error_excel=True
##            error_audio=True
##            error_time=True
##            
##            if not self.Excel_filepath:
##                self.le.setText("please Select Excel File First!!!")
##                error_excel= False
##                print("Error file:",error_excel)
##            
##            if not self.Input_audio_path:
##                self.sel_audio.setText("please Select Audio File!!!")
##                error_audio= False
##                print("Error audio:",error_audio)
##            try:
##                if self.Input_time_Perslide == 0:
##                    print(self.Input_time_Perslide,"time")
##                    self.sel_msg.setText("Enter Duration!!!")
##                    error_time=False
##                    print("error_time:",error_time)
##            except Exception as e:
##                self.sel_msg.setText("Error Occured Enter Correct Time between 0 _ 20 ")
##
##            if(error_excel==True and error_audio==True and error_time==True):
##                wb = xlrd.open_workbook(self.Excel_filepath) 
##                sheet = wb.sheet_by_index(0)
##                arr=[]
##                rows = sheet.nrows
##                columns = sheet.ncols
##                self.No_of_Row=columns
##                self.Input_time_Perslide=self.sel_time.text()
##                os.chdir(self.Input_audio_path)
##                path=self.Input_audio_path   #self.Input_audio_path_first
##                limit=0;
##                time=int(self.Input_time_Perslide)
##                print("time:",time)
##                col=columns
##                limit=(time*col)+(time*2)
##                print("LIMIT:",limit)
##                count_audio=0;
##                if rows ==0:
##                    print(self.No_of_Row)
##                    self.sel_msg.setText("Your Excel File is Empty Enter a Valid Excel File !!!!!! ")
##                else:
##                    audfile=[]
##                    for root, dirs, files in os.walk(path):
##                        for file in files:
##                            p=os.path.join(root,file)
##                            if p.endswith(".mp3"):
##                                audfile.append(os.path.abspath(p))
##
##                    size=0;
##                    for i in audfile:
##                        audio_background = AudioFileClip(i)
##                        print("Duration of Audio:",audio_background.duration)
##                        aud_duration=audio_background.duration
##                        print("aud_duration:",aud_duration)
##                        size=int(aud_duration)
##                        audio_background.close()
##                        if(limit<size):
##                            self.Error_Audio_Folder=True
##                            count_audio=count_audio+1;
##                            break;
##                if(self.Error_Audio_Folder==False):
##                    self.sel_msg.setText("Error Occured  Enter a Audio files which has size greater than  "+str(limit)+" seconds")
####                if(self.Error_Audio_Folder==True):
####                    self.sel_msg.setText(str(count_audio)+" Audio Files are Valid for process !!!")
##                
##        except Exception as e:
##            val=str(e)
##            #print("Exception Occured !!!",val)
##            self.sel_msg.setText("Error Occured !!! Enter Fields correctly ")    
    
    def validate_input(self):
        print("validateInput")
        try:
            self.Input_time_Perslide=self.sel_time.text()
            check=self.RepresentsInt(self.Input_time_Perslide)
            if(check==True):
                self.Input_time_Perslide=self.sel_time.text()
            else:
                self.Input_time_Perslide=""
            #print("error_time:",check)
            #self.Input_time_Perslide=self.sel_time.text()
            error_excel=True
            error_img=True
            error_intro=True
            error_outro=True
            error_audio=True
            error_dest=True
            error_time=True
            print("Input_time_Perslide",self.Input_time_Perslide)
            if not self.Excel_filepath:
                self.le.setText("please Select Excel File First!!!")
                error_excel= False
                #print("Error file:",error_excel)
            if not self.Image_folder_path:
                self.sel_folder.setText("please Select Images Folder!!!")
                error_img= False
                #print("Error folder img:",error_img)
            if not self.Input_intro_path:
                self.sel_intro.setText("please Select Introduction First!!!")
                error_intro= False
                #print("Error Input:",error_intro)
            if not self.Input_outro_path:
                self.sel_outro.setText("please Select Outro  First!!!")
                error_outro= False
                #print("Error outro:",error_outro)
            if not self.Input_audio_path:
                self.sel_audio.setText("please Select Audio File!!!")
                error_audio= False
                #print("Error audio:",error_audio)
            if not self.Input_dest_path:
                self.sel_dest.setText("please Select Destination Folder!!!")
                error_dest= False
                #print("Error dest:",error_dest)
            if not self.Input_time_Perslide:
                self.sel_time.setText("?")
                self.sel_msg.setText("Enter Duration!!!")
                error_time=False
                
            if(check==True and error_excel==True and error_img==True and error_intro==True and error_outro==True and error_audio==True and error_dest==True and error_time==True):
                self.sel_msg.setText("process start !!!!!")
                self.Read_Excel_File()
                   # print("start Now ______")
        except Exception as e:
            #print(e)
            msg=str(e)
            self.sel_msg.setText(msg)
            
    def onCountChanged(self, value):
        self.progress.setValue(value)
        message="process is  : "+str(value)+" %  done Wait !!!!"
        self.sel_msg.setText(message)
   
        # Processing of backend starts here
    def Read_Excel_File(self):
        #print("END_read")

        
        #self.btn_start.setEnabled(False)
        self.myThread = Worker(self.Excel_filepath,self.Image_folder_path,self.Input_intro_path,self.Input_outro_path,self.Input_audio_path,self.Input_dest_path,self.Input_text_color,self.Input_back_color,self.Input_time_Perslide,self.Aud,self.No_of_Row)   
        self.myThread.countChanged.connect(self.onCountChanged)
        self.myThread.start_activation.connect(self.onstart_activation)
        self.myThread.setError.connect(self.onsetError)
        self.sig.connect(self.myThread.on_stopprocess)
        #self.sigstop.connect(self.myThread.on_startprocess)
        self.myThread.start()
        #print("end of thread1")

    def onstart_activation(self,value):
        print("start Func hit --------------------------------------value :",value)
        if value==True:
            print("True if hit _--------------------------------------------------")
            self.btn_start.setEnabled(True)
            #if self.count_error==False:
            self.sel_msg.setText("Process is Finished Now !!!")
            self.btn_cancel.setEnabled(False)
        if value==False:
             self.btn_start.setEnabled(False)
           #  print("start button is de-active now ")
             #self.sel_msg.setText("Process is Finsihed")
             self.btn_cancel.setEnabled(True)
    def onstop_work(self,value):
        self.sig.emit(True)
        self.sel_msg.setText("Process is going to stop Wait !!!!")
        #self.btn_start.setEnabled(True)
        self.btn_cancel.setEnabled(False)
        #self.sel_msg.setText("Process is Cancelled")

    def onsetError(self,value):
        self.count_error=True
        #print("Error emit",value)
        msg="Error Occured "+value+"  !!!!!!"
        self.sel_msg.setText(msg)
        self.btn_start.setEnabled(True)
            
def window():
    app = QApplication(sys.argv)
    win = MyWindow()
    win.show()
    sys.exit(app.exec_())

window()
