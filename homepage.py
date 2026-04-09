import streamlit

streamlit.title('Be Bio Python Server Homepage')
streamlit.header('Instructions')
streamlit.text('Welcome! To get started, click the application you want to use from the list of apps on the left hand side of the page. Once the new page loads, fill in the required information, give the script a little time to run, and then you can download your results!')
streamlit.image("Logo.png")
pg = streamlit.navigation([
	streamlit.Page('pages/ddPCRAutomation_17OCT25Updates_4 (1).py', title = 'HDR ddPCR Automation', icon = ":material/open_in_new:", url_path = 'HDRddPCR')
	])

pg.run()




