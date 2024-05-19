import streamlit as st
import streamlit_authenticator as stauth # pip install streamlit-authenticator
import yaml #pip install yaml
from yaml.loader import SafeLoader
import os
# pip install pipreqs -->再去 terminal 入pipreqs
#
# emojis: https://www.webfx.com/tools/emoji-cheat-sheet/
st.set_page_config(page_title="streamlit Dashboard", page_icon=":bar_chart:", layout="wide")#
#Move the title higher
st.markdown('<style>div.block-container{padding-top:1rem;}</style>',unsafe_allow_html=True)

import json #不用pip install, show animation
import streamlit.components.v1 as com #用frame show animation
from streamlit_lottie import st_lottie  #pip install streamlit-lottie

hide_bar= """
    <style>
    [data-testid="stSidebar"][aria-expanded="true"] > div:first-child {
        visibility:hidden;
        width: 0px;
    }
    [data-testid="stSidebar"][aria-expanded="false"] > div:first-child {
        visibility:hidden;
    }
    </style>
"""

# load hashed passwords+ load yaml preset account& pw info
hashed_passwords = stauth.Hasher(['smt888', 'smt888']).generate() #要load一次generate yaml 密碼再copy & paste去replace yaml檔的原有密碼

#用來放yaml檔的路徑
#os.chdir(r"/Users/arthurchan/Documents/PythonProject")


with open('config.yaml', encoding="utf-8") as file:
    config = yaml.load(file, Loader=SafeLoader)

# --- USER AUTHENTICATION ---
authenticator = stauth.Authenticate(
    config['credentials'],
    config['cookie']['name'],
    config['cookie']['key'],
    config['cookie']['expiry_days'],
    config['preauthorized']
)

name, authentication_status, username = authenticator.login(':closed_lock_with_key: SMT Sales Dashboard Login', 'main')

if authentication_status == False:
    st.error('Username/password is incorrect')
    st.markdown(hide_bar, unsafe_allow_html=True)
    #用iframe加入小框架json animation- <iframe src="https://lottie.host/embed/c3e4ccd6-6f7d-4cd9-b172-f8b997d10226/yjXddFZ51Y.json"></iframe> 去lottie copy iframe code
    com.iframe("https://lottie.host/embed/c3e4ccd6-6f7d-4cd9-b172-f8b997d10226/yjXddFZ51Y.json",height=200, width= 200)

elif authentication_status == None:
    st.warning('Please enter your username and password')
    st.markdown(hide_bar, unsafe_allow_html=True)
    #用iframe加入小框架json animation- <iframe src="https://lottie.host/embed/00483453-023f-40b1-8ad6-ea4ee512c696/oUzzb7E18e.json"></iframe> 去lottie copy iframe code
    com.iframe("https://lottie.host/embed/00483453-023f-40b1-8ad6-ea4ee512c696/oUzzb7E18e.json",height=200, width= 200)
    

elif authentication_status:
     st.success('Login Success!')
      

###############################################

     left_column, middle_column, right_column1, right_column2 = st.columns(4)
     with left_column:
        # Welcome Title
        st.title(f'Welcome *{name}*! :sunglasses:')
        st.header(":bulb: Introduction")
        st.divider()
        ###about ....
        st.subheader(":one: Invoice Summary:")
        st.write("Based on 5 years of Invoice Data, multi-perspective business analysis on 4 sections: **OVERALL**, **REGION**, **BRAND** and **CUSTOMER**")
        #用iframe加入小框架json animation- <iframe src="https://lottie.host/embed/2500fe42-7188-45a0-ad1f-5e0afc6b6026/5mu4rRYTtA.json"></iframe> 去lottie copy iframe code
        com.iframe("https://lottie.host/embed/2500fe42-7188-45a0-ad1f-5e0afc6b6026/5mu4rRYTtA.json",height=300, width= 500)
        st.divider()
         
        st.subheader(":two: Contract Summary:")
        st.write("Based on 5 years of Contract Data, multi-perspective business analysis on 4 sections: **OVERALL**, **REGION**, **BRAND** and **CUSTOMER**")
        #用iframe加入小框架json animation- <iframe src="https://lottie.host/embed/f2b3228f-61e1-4c02-b1f7-ed59315ac5df/f5MuluZ3N9.json"></iframe> 去lottie copy iframe code
        com.iframe("https://lottie.host/embed/f2b3228f-61e1-4c02-b1f7-ed59315ac5df/f5MuluZ3N9.json",height=300, width= 500)
        st.divider()
         
        st.subheader(":three: Mounter Import data of China:")
        st.write("Based on 5 years of Mounter Import data of China, business analysis on the mounter sales trend")
        #用iframe加入小框架json animation- <iframe src="https://lottie.host/embed/0b2ec926-dfed-47d6-81f9-37828e796a74/CEVkcXN1pz.json"></iframe> 去lottie copy iframe code
        com.iframe("https://lottie.host/embed/0b2ec926-dfed-47d6-81f9-37828e796a74/CEVkcXN1pz.json",height=300, width= 500)
        st.divider()
         
        st.subheader(":four: Sales Projection:")
        st.write("Based on **customized growth rate** to build a sales projection for reference")
        #用iframe加入小框架json animation- <iframe src="https://lottie.host/embed/ec15cea4-eaff-4e4f-bef5-cac85cbe7b5e/BhbRcST6Me.json"></iframe> 去lottie copy iframe code
        com.iframe("https://lottie.host/embed/ec15cea4-eaff-4e4f-bef5-cac85cbe7b5e/BhbRcST6Me.json",height=300, width= 500)
        st.divider()
        
     with middle_column:
     #Animation- lab table
        #用iframe加入小框架json animation- <iframe src="https://lottie.host/embed/db52db3f-d63e-4796-87e4-9e6e33058186/LFuU45fNO6.json"></iframe> 去lottie copy iframe code
        com.iframe("https://lottie.host/embed/db52db3f-d63e-4796-87e4-9e6e33058186/LFuU45fNO6.json",height=600, width= 300)
     with right_column1:
        st.title(":coffee: User Guide:")
        st.header(":eight_pointed_black_star: Attention")
        st.divider()
        st.subheader(":white_check_mark: Download/ Full Screen/ Zoom in:")
        st.write("All charts and tables in each section are avaiable to **zoom in**, view in **full screen** and **quick download**")
         #用iframe加入小框架json animation- <iframe src="https://lottie.host/embed/4c312e0b-3758-4901-bcba-d855af2651a3/T8qfArmT2l.json"></iframe> 去lottie copy iframe code
        com.iframe("https://lottie.host/embed/4c312e0b-3758-4901-bcba-d855af2651a3/T8qfArmT2l.json",height=300, width= 500)
        st.divider()

        st.subheader(":white_check_mark: Security:")
        st.write("User does not need to log out when exit, it will log out **automatically** when closing the webpage")
        #用iframe加入小框架json animation- <iframe src="https://lottie.host/embed/fb60e365-6c43-46e9-80cc-39bd8ec30d54/Tme9m9fYCg.json"></iframe> 去lottie copy iframe code
        com.iframe("https://lottie.host/embed/fb60e365-6c43-46e9-80cc-39bd8ec30d54/Tme9m9fYCg.json",height=300, width= 500)

    #Animation- user guide
     with right_column2:
     #Animation- presentation
        #用iframe加入小框架json animation- <iframe src="https://lottie.host/embed/716da532-a8b9-42f9-b5b5-dc375ae80148/HMAcHQBE6q.json"></iframe> 去lottie copy iframe code
        com.iframe("https://lottie.host/embed/716da532-a8b9-42f9-b5b5-dc375ae80148/HMAcHQBE6q.json",height=500, width= 400)

st.sidebar.success(":point_up_2: Select a page above.")
     ###---- HIDE STREAMLIT STYLE ----
hide_st_style = """
                <style>
                #MainMenu {visibility: hidden;}
                footer {visibility: hidden;}
                header {visibility: hidden;}
                </style>
                """
st.markdown(hide_st_style, unsafe_allow_html=True)
     
authenticator.logout("Logout", "sidebar",key="unique_key")



