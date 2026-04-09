# NDS Capture RPA Bot

Early Python-based RPA project that automates capture workflows using Excel, UI automation, and email notifications.

---

## Overview

This project automates manual capture processes by reading Excel files, interacting with a desktop system, and processing debit and credit transactions.

---

## Features

- Excel-driven processing
- UI automation using PyAutoGUI
- Email notifications
- Splunk logging

---

## Configuration

This project uses a configuration file for environment-specific values.

A sample file `config.example.ini` is included.

To run locally:
1. Copy `config.example.ini`
2. Rename it to `config.ini`
3. Update values with your local setup

⚠️ Do not commit real credentials or sensitive data.

##  How It Works

1. Read Excel files from input folder  
2. Clean and prepare transaction data  
3. Log into the desktop system  
4. Perform debit and credit capture  
5. Move processed files to success or exception folders  
6. Send email notifications  
7. Log results to Splunk  

---



## 🛠️ Technologies Used

- Python  
- PyAutoGUI  
- pandas  
- openpyxl  
- requests  

---

##  What I Learned

- UI automation for legacy systems  
- Handling Excel-based workflows  
- Structuring automation scripts  
- Logging and monitoring  

---

## ️ Notes

This is a cleaned version of a real-world automation project.  
Sensitive data and internal configurations have been removed.