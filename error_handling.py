import logging

def handle_error(error, message):
    logging.error(f"{message}: {error}")
    print(f"An error occurred: {message}. Please check the log for more details.")

def handle_file_error(error):
    handle_error(error, "Error while handling file")

def handle_data_ingestion_error(error):
    handle_error(error, "Error during data ingestion")

def handle_data_parsing_error(error):
    handle_error(error, "Error during data parsing")

def handle_mrr_per_client_error(error):
    handle_error(error, "Error while creating MRR Per Client table")

def handle_mrr_status_error(error):
    handle_error(error, "Error while creating MRR Status table")

def handle_gui_error(error):
    handle_error(error, "Error in the GUI")