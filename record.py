import time
import threading
from moviepy.video.io.VideoFileClip import VideoFileClip
from moviepy.audio.io.AudioFileClip import AudioFileClip
import os
import subprocess
import wave
import pyautogui
import numpy as np
from pathlib import Path
import logging
import cv2
from SMWinservice import SMWinservice
import pyaudio
from win32api import GetSystemMetrics
log_path="C:\\RecordProgram\\logs"
video_path="C:\\RecordProgram\\video"
audio_clip_path="C:\\RecordProgram\\video\\output.wav"
clip_path="C:\\RecordProgram\\video\\output.avi"
dir="C:\\RecordProgram\\"
host_name="Stereo Karışımı"
w=GetSystemMetrics(0)
h=GetSystemMetrics(1)

def create_path(path):
    if os.path.exists(path):
        pass
    else:
        os.mkdir(path)
        
        if path=="C:\\RecordProgram":
            file = open(dir+"num.txt", "w") 
            file.write("0") 
            file.close() 
create_path(dir[:-1])
create_path(log_path)

logger=logging.getLogger(__name__)
formatter=logging.Formatter("%(lineno)d:%(asctime)s:%(levelname)s")
fileHandler=logging.FileHandler(log_path+"\\record.log")
fileHandler.setFormatter(formatter)
logger.addHandler(fileHandler)

class RecorderPython(SMWinservice):
    
    
    _svc_name_ = "ARecordService1"
    _svc_display_name_ = "ARecordService1"
    _svc_description_ = "That's a great winservice! :)"
    
    def stop(self):
        self.isrunning=False
    def start(self):
        self.isrunning = True
    def rec(self):
       
        self.create_video_file()
        if self.read_text()==1:
            try:
                self.remove_file()
                
                writer = cv2.VideoWriter(clip_path, cv2.VideoWriter_fourcc(*"MJPG"), 13, (w, h))
                while True:
                  
                        
                        frame = pyautogui.screenshot()
                        frame = cv2.cvtColor(np.array(frame), cv2.COLOR_BGR2RGB)
                        
                        writer.write(frame)
                        if self.read_text()==0:
                            break

                writer.release()
                cv2.destroyAllWindows()
                self.setDuration()
                
            except Exception as e:
                logger.error(str(e))

    def voice_record(self):
        if self.read_text()==1:
            try:
                
                self.remove_file()
                CHUNK = 1024
                FORMAT = pyaudio.paInt16
                CHANNELS = 2
                RATE = 44100
                host_index=[]
                WAVE_OUTPUT_FILENAME = audio_clip_path
                p = pyaudio.PyAudio()
                [host_index.append(p.get_device_info_by_index(i)["index"]) if host_name in p.get_device_info_by_index(i)["name"]  else None for i in range(p.get_device_count())]
                stream = p.open(format=FORMAT,
                                channels=CHANNELS,
                                rate=RATE,
                                input=True,
                                frames_per_buffer=CHUNK,
                                input_device_index=host_index[0])
                frames = []
                while True:
                    if self.read_text()==1:
                        data = stream.read(CHUNK)
                        frames.append(data)
                        if self.read_text() == 0:
                            break
                        wf = wave.open(WAVE_OUTPUT_FILENAME, 'wb')
                        wf.setnchannels(CHANNELS)
                        wf.setsampwidth(p.get_sample_size(FORMAT))
                        wf.setframerate(RATE)
                        wf.writeframes(b''.join(frames))
                        wf.close()
                stream.stop_stream()
                stream.close()
                p.terminate()
            except Exception as e:
                logger.error(str(e))

    def thread(self):
        try:
            self.t1=threading.Thread(target=lambda: self.voice_record())
            self.t2=threading.Thread(target=lambda: self.rec())
            self.t1.start()
            self.t2.start()
            self.t1.join()
            self.t2.join()
            
        except Exception as e:
            logger.error(str(e))
        

    def setDuration(self):
        try:
            audioclip = AudioFileClip(audio_clip_path)
            clip = VideoFileClip(clip_path)
            
            self.file_exist = 0
            while True:

                if os.path.exists(video_path+"\\"+str(self.file_exist) + ".avi") == 0:
                    durations1 = round(int(audioclip.duration) / int(clip.duration), 4)
                    file_loc = clip_path
                    output_loc = dir+"video\\{}.avi".format(self.file_exist)
                    set_duration = subprocess.Popen(f"ffmpeg -y -i {file_loc} -vf \"setpts={durations1}*PTS\" -r 24 {output_loc}",
                    stdout=subprocess.PIPE, shell=True)
                    
                    set_duration.communicate()
                    set_duration.kill()
                    break
                            
                else:
                    self.file_exist = self.file_exist + 1
            
        except Exception as e:
            logger.error(str(e))
    def combine(self) :
        try:
            time.sleep(3)
            self.setDuration()
            import ffmpeg
            video_stream = ffmpeg.input(dir+'video\\{}.avi'.format(self.file_exist))
            combine_wav =subprocess.Popen('ffmpeg -y -i ./video/output.wav -i ./video/microphone.wav -filter_complex "[0:0][1:0] amix=inputs=2:duration=longest" -c:a libmp3lame ./video/output5.wav',
            stdout=subprocess.PIPE, shell=True)
            combine_wav.communicate()
            combine_wav.kill()
            audio_stream = ffmpeg.input(dir+'video\\output5.wav')
            t = str(time.strftime('%X').replace(':', '.').replace("/", ".").replace("\\", "."))
            mp4_name = dir+"video\\{}.mp4".format(t)
            ffmpeg.output(audio_stream, video_stream, mp4_name).run()
            self.remove_file()
            
        except Exception as e:
            
            logger.error(str(e))


    def remove_file(self):
        try:
            filenames = []
            for i in ["*.avi", "*.wav"]:
                filenames.extend([x for x in Path(dir+"video").glob("{}".format(i))])
            for j in range(len(filenames)):
                try:
                    os.remove(str(filenames[j]))
                except :
                    continue
        except Exception as e:
            pass

    def create_video_file(self):
        try:
            if os.path.exists(video_path):
                pass
            else:
                os.mkdir(video_path)
        except Exception as e:
            logger.error(str(e))

    def main(self):        
            if self.read_text()==1:
                
                self.thread()
            
                  
    def read_text(self):
        num=""
        
        with open(dir+"num.txt","r",encoding="utf-8") as f:
            num=f.read()
            
        if num=="0":
            return 0
        else :
            return 1

if __name__=='__main__':
    RecorderPython.parse_command_line()