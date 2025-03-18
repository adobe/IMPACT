# Copyright 2025 Adobe. All rights reserved.
# This file is licensed to you under the Apache License, Version 2.0 (the "License");
# you may not use this file except in compliance with the License. You may obtain a copy
# of the License at http://www.apache.org/licenses/LICENSE-2.0

# Unless required by applicable law or agreed to in writing, software distributed under
# the License is distributed on an "AS IS" BASIS, WITHOUT WARRANTIES OR REPRESENTATIONS
# OF ANY KIND, either express or implied. See the License for the specific language
# governing permissions and limitations under the License.

import paramiko
import os,socket

def new_copy_folder_via_scp(local,remote):
    host=''
    port=22
    username=''
    password=''
    transport = paramiko.Transport((host,port))
    transport.connect(username = username,password = password)
    sftp = paramiko.SFTPClient.from_transport(transport)
    try:
        folderName=local.split('/')[-1]
        print(folderName)
        newFolder=os.path.join(remote+"/"+folderName)
        print(newFolder)
        print(sftp.mkdir(newFolder))
        remote=newFolder
        print(remote)
        if os.path.isdir(local):
            for f in os.listdir(local):
                print("F : ",f)
                if not os.path.isdir((os.path.join(local+"/"+f))):
                    print("uploading:",os.path.join(local+"/"+f),os.path.join(remote+"/"+f))
                    sftp.put(os.path.join(local+"/"+f),os.path.join(remote+"/"+f))
                    continue
                print('mkdir: ',os.path.join(remote+"/"+f),"returning: ", sftp.mkdir(os.path.join(remote+"/"+f)))
                for d in os.listdir(os.path.join(local+"/"+f)):
                    if(os.path.isdir(os.path.join(local+"/"+f+"/"+d))):
                        new_copy_folder_via_scp(host,port,username,password,local+"/"+f+"/",remote+"/"+f+'/')
                        continue
                    print("uploading:",os.path.join(local+"/"+f+"/"+d),os.path.join(remote+"/"+f+'/'+d))
                    sftp.put(os.path.join(local+"/"+f+"/"+d),os.path.join(remote+"/"+f+'/'+d))
        else:
            sftp.put(local,remote)
    except Exception as e:
        print('exception:',e)
    transport.close()
    print("Upload Completed !")