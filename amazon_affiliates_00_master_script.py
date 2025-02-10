#!/usr/bin/env python
# coding: utf-8

# In[4]:

def amazon_affiliates_00_master_script():
    from amazon_affiliates_01_download_report import amazon_affiliates_01_download_report
    from amazon_affiliates_02_processing_job import amazon_affiliates_02_processing_job

    amazon_affiliates_01_download_report()
    amazon_affiliates_02_processing_job()


if __name__ == "__main__":
    amazon_affiliates_00_master_script()
