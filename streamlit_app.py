import streamlit as st
import json, time
import requests
import re,os
from io import BytesIO
from bs4 import BeautifulSoup, Tag
from bs4.element import NavigableString
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH

import base64, os
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
header_img_path = BASE_DIR / "static" / "SIDBI_LOGO.png"
# header_img_path = "SIDBI_LOGO.png"

# safely embed logo

API_BASE = os.getenv("API_BASE", "http://127.0.0.1:8000")
# ================== CONFIG ==================
st.set_page_config(page_title="AI Document Creator", layout="wide")
img_tag = ""
try:
    if header_img_path.exists():
        with open(header_img_path, "rb") as f:
            img_b64 = base64.b64encode(f.read()).decode("utf-8")
        img_tag = f"<img src='data:image/png;base64,{img_b64}' alt='logo' style='height:64px; max-width:250px; object-fit:contain; display:block;'/>"
except:
    img_tag = ""
# # For main header in page

html2 = f"""
<div class ="">
<div class="header-wrap">
  <div class="header-logo">
    {img_tag}
    
  </div>
  <div class="">
  <h2 style='text-align: center; color: black; padding: 1px 0 0 0'>AI Document Creator</h2>
   <p style='margin-bottom:0'>
            Upload your source PDF, then let AI summarize and auto-populate a ready-to-download document.
     </p>
  </div>
 
  </div>
  
   </div>


"""  
st.markdown(html2, unsafe_allow_html=True)


html = f"""
<style>
.st-emotion-cache-zy6yx3 {{
    padding: 10px;

}}

.header-wrap {{

    display: flex;
    justify-content: left;
    align-items: center;
    column-gap: 15%;
    box-shadow: 1px 2px 5px #bdbcbc;
    
}}

.header-container {{
    box-shadow: 1px 2px 5px #bdbcbc;
}}

.stAppHeader{{
display: none !important;
}}
</style>


"""
st.markdown(html, unsafe_allow_html=True)


# ---------- CONFIG: Backend URLs ----------
LLM_API_URL = "http://172.30.1.200:9000/extract"  # LLM API
# LLM_API_URL = "https://file-extractordev.sidbi.in/fill-docx-service/extract"
DOCX_BACKEND_URL = f"{API_BASE}/docx-to-json"
PDF_BACKEND_URL  = f"{API_BASE}/pdf-to-text"
FILL_HTML_URL    = f"{API_BASE}/fill-html"
HTML_TO_DOCX_URL = f"{API_BASE}/html-to-docx"



# Hardcoded template values - adjust as needed for your backend
# IM_TEMPLATE_SCHEMA = [{'placeholders': {'Name_of_Issuer_Company': '', 'MSME_status/_Udyam_Status': '', 'Date_of_Incorporation': '', 'Commencement_of_Operations': '', 'Address_of_Registered_Office': '', 'Address_of_Operating_facilities': ''}}, {'total_issue_size_price_in_crore':''},{'Brief_Background_of_company_and_promoters': ''}, {'Details_about_DRHP_filing_date_approval_date': ''}, {'Details_about_Exchange_on_which_company_is_going_to_be_listed_Merchant_banker_and_Registrar.': ''}, {'placeholders': {'IPO_Date': '', 'Listing_Date': '', 'Face_Value': '', 'Issue_Price_Band': '', 'Lot_Size': '', 'Sale_Type': '', 'Issue_size': '', 'Issue_Type': '', 'Listing_At': '', 'Share_Holding_Pre_Issue': '', 'Share_Holding_Post_Issue': ''}}, {'placeholders': {'QIB_Shares_Offered_Shares_Offered': '', 'Retail_Shares_Offered_Shares_Offered': '', 'NII_Shares_Offered_Shares_Offered': ''}}, {'placeholders': {'Assets_31_March_2025': '', 'Assets_31_Mar_2024': '', 'Assets_31_Mar_2023': '', 'Total_Income_31_March_2025': '', 'Total_Income_31_Mar_2024': '', 'Total_Income_31_Mar_2023': '', 'Profit_After_Tax_31_March_2025': '', 'Profit_After_Tax_31_Mar_2024': '', 'Profit_After_Tax_31_Mar_2023': '', 'Net_Worth_31_March_2025': '', 'Net_Worth_31_Mar_2024': '', 'Net_Worth_31_Mar_2023': '', 'Reserves_and_Surplus_31_March_2025': '', 'Reserves_and_Surplus_31_Mar_2024': '', 'Reserves_and_Surplus_31_Mar_2023': '', 'Total_Borrowing_31_March_2025': '', 'Total_Borrowing_31_Mar_2024': '', 'Total_Borrowing_31_Mar_2023': ''}}, {'Brief_about_company’s_projections_profitability_etc': ''}, {'placeholders': {'Promoter_group_Shareholding_Pattern_Pre-IPO_Nos': '', 'Promoter_group_Shareholding_Pattern_Pre-IPO_%': '', 'Promoter_group_Shareholding_Pattern_Post_IPO_Nos': '', 'Promoter_group_Shareholding_Pattern_Post_IPO_%': '', 'Subtotal_(A)_Shareholding_Pattern_Pre-IPO_Nos': '', 'Subtotal_(A)_Shareholding_Pattern_Pre-IPO_%': '', 'Subtotal_(A)_Shareholding_Pattern_Post_IPO_Nos': '', 'Subtotal_(A)_Shareholding_Pattern_Post_IPO_%': '', 'AIF/Individuals/corporates_Shareholding_Pattern_Pre-IPO_Nos': '', 'AIF/Individuals/corporates_Shareholding_Pattern_Pre-IPO_%': '', 'AIF/Individuals/corporates_Shareholding_Pattern_Post_IPO_Nos': '', 'AIF/Individuals/corporates_Shareholding_Pattern_Post_IPO_%': '', 'IPO_shareholders_Shareholding_Pattern_Pre-IPO_Nos': '', 'IPO_shareholders_Shareholding_Pattern_Pre-IPO_%': '', 'IPO_shareholders_Shareholding_Pattern_Post_IPO_Nos': '', 'IPO_shareholders_Shareholding_Pattern_Post_IPO_%': '', 'Subtotal_(B)_Shareholding_Pattern_Pre-IPO_Nos': '', 'Subtotal_(B)_Shareholding_Pattern_Pre-IPO_%': '', 'Subtotal_(B)_Shareholding_Pattern_Post_IPO_Nos': '', 'Subtotal_(B)_Shareholding_Pattern_Post_IPO_%': '', 'Total_(A+B)_Shareholding_Pattern_Pre-IPO_Nos': '', 'Total_(A+B)_Shareholding_Pattern_Pre-IPO_%': '', 'Total_(A+B)_Shareholding_Pattern_Post_IPO_Nos': '', 'Total_(A+B)_Shareholding_Pattern_Post_IPO_Percentage': ''}}, {'placeholders': {'DRHP_should_have_been_approved_by_any_Stock_Exchange_/_SEBI._Compliance': '', 'Should_have_availed_at_least_one_Debt_products_from_Banks/FIs_and_servicing_it_for_at_least_3_years_prior_to_the_date_of_filing_of_DRHP._and_/_or_should_have_availed_funding_from_any_SEBI_registered_AIFs_at_least_1_year_prior_to_the_date_of_filing_of_DRHP._Compliance': '', 'Net_worth_>_₹_20_crore_Compliance': '', 'EBIDTA_margin_-_₹2.00_cr._or_>_5%_(whichever_is_more)_for_at_least_two_out_of_the_three_most_recent_financial_years_Compliance': '', 'Total_Income_>_₹50_crore_Compliance': '', 'Net_Tangible_Assets_>_₹_2.50_crore_Compliance': '', 'Profitability_(total_PBT_of_3_years)_>_₹_10_crore_Compliance': '', 'Exposure_cap_per_MSME:_Minimum:_₹_1_crore_Maximum:_Up_to_₹_20_crore._Note:_Anchor_investment_shall_not_exceed_50%_of_the_Anchor_portion_of_the_Issue_or_10%_of_the_post_issue_paid-up_capital_of_the_issuer_company_whichever_is_lower._Compliance': '', 'Minimum_Promoters_Shareholding_(Post_Issue)-_60%_In_case_where_the_Company_has_received_any_investment_from_a_SEBI_registered_AIF_the_Investment_Committee_may_take_a_suitable_view_on_the_minimum_level_of_Promoter_Shareholding_below_60%_since_the_AIF_would_want_to_partially_/_fully_exit_in_the_proposed_IPO_/_Offer_for_Sale_(OFS)._Compliance': '', 'Complied_with_Minimum_₹_2_Cr._Bid_criteria_of_SEBI_Compliance': '', 'Complied_with_RBI’s_guideline_limiting_post-issue_shareholding_exposure_to_a_maximum_of_10%._Compliance': ''}}, {'Brief_about_industry': ''}, {'placeholders': {'1_Items': '', '1_Amount_(₹_crore)': '', '1_%': '', '2_Items': '', '2_Amount_(₹_crore)': '', '2_%': '', '3_Items': '', '3_Amount_(₹_crore)': '', '3_%': '', '4_Items': '', '4_Amount_(₹_crore)': '', '4_%': '', 'Total_Items': '', 'Total_Amount_(₹_crore)': '', 'Total_%': ''}}, {'Brief_about_basis_of_valuation': ''}, {'placeholders': {'PAT_(in_lakh)_Mar-25': '', 'PAT_(in_lakh)_Mar-24': '', 'PAT_(in_lakh)_Mar-23': '', 'EPS_Mar-25': '', 'EPS_Mar-24': '', 'EPS_Mar-23': '', 'EPS_after_adjusting_Bonus_Mar-25': '', 'EPS_after_adjusting_Bonus_Mar-24': '', 'EPS_after_adjusting_Bonus_Mar-23': '', 'EPS_after_Proposed_Issue_Mar-25': '', 'EPS_after_Proposed_Issue_Mar-24': '', 'EPS_after_Proposed_Issue_Mar-23': '', 'Weight_Mar-25': '', 'Weight_Mar-24': '', 'Weight_Mar-23': '', 'EPS_Product_Mar-25': '', 'EPS_Product_Mar-24': '', 'EPS_Product_Mar-23': '', 'PE_Ratio_(2025_EPS)_Mar-25': '', 'PE_Ratio_(2025_EPS)_Mar-24': '', 'PE_Ratio_(2025_EPS)_Mar-23': '', 'PE_Ratio_Weighted_Average_EPS_Mar-25': '', 'PE_Ratio_Weighted_Average_EPS_Mar-24': '', 'PE_Ratio_Weighted_Average_EPS_Mar-23': '', 'PE_Ratio_Post_issue_based_on_FY_2025_Earning_Mar-25': '', 'PE_Ratio_Post_issue_based_on_FY_2025_Earning_Mar-24': '', 'PE_Ratio_Post_issue_based_on_FY_2025_Earning_Mar-23': '', 'Price_(Rs)_Mar-25': '', 'Price_(Rs)_Mar-24': '', 'Price_(Rs)_Mar-23': '', 'Issue_Size_(Rs_Cr)_-_Public_Mar-25': '', 'Issue_Size_(Rs_Cr)_-_Public_Mar-24': '', 'Issue_Size_(Rs_Cr)_-_Public_Mar-23': '', 'Post_Issue_Valuation_of_Mar-25': '', 'Post_Issue_Valuation_of_Mar-24': '', 'Post_Issue_Valuation_of_Mar-23': '', 'Company_(Rs_Cr)_Mar-25': '', 'Company_(Rs_Cr)_Mar-24': '', 'Company_(Rs_Cr)_Mar-23': '', 'Pre_Issue_PBV_Mar-25': '', 'Pre_Issue_PBV_Mar-24': '', 'Pre_Issue_PBV_Mar-23': '', 'Post_Issue_Book_Value_Mar-25': '', 'Post_Issue_Book_Value_Mar-24': '', 'Post_Issue_Book_Value_Mar-23': '', 'Post_Issue_PBV_Mar-25': '', 'Post_Issue_PBV_Mar-24': '', 'Post_Issue_PBV_Mar-23': '', 'Post_Issue_Networth_(cr)_Mar-25': '', 'Post_Issue_Networth_(cr)_Mar-24': '', 'Post_Issue_Networth_(cr)_Mar-23': ''}}, {'placeholders': {'Company_Name': '', 'FV': '', 'Net_worth': '', 'Sales': '', 'PAT': '', 'EPS': '', 'Issue_Price': '', 'P/E': '', 'B.V_Pre-Issue(`)': '', 'PBV': '', 'RONW_(%)': '', '(`)': '', '(x)': ''}},{'total_issue_size_price_in_crore':''},{'Number_of_shares': ''}, {'Number_of_lots': ''}, {'Number_of_shares': ''}, {'about_promoters_experience': ''}, {'about_consistent_in_profitability': ''}, {'about_entry_barriers': ''}, {'any_other_rationale': ''}, {'About_Lead_and_Co-lead_merchant_bankers_their_past_track_record': ''}, {'Brief_about_Market_Maker_and_their_track_record': ''}, {'On_the_basis_of_Credibility_of_issuer_Mechant_banker_potential_of_returns_Good_investors_backing_etc': ''}, {'Date': ''},{'placeholders': {'Compliance_of_KYC_guidelines': '', 'KYC_Risk_Score': '', 'Verification_of_defaulters’_lists_Watchout_Investor_website_Income_Tax_defaulter_Search_etc.': '', 'Verification_of_CIBIL_reports_For_Company_and_Promoters': '', 'JOCATA_Score': '', 'Market_Inquiries/_Feedback_from_references': '', 'Connected_Lending': '', 'Pending_court_cases_initiated_by_other_banks/_FIs_against_the_company_promoters/director_etc_if_any': '', 'Pending_court_cases_initiated_by_third_parties_(Debtors/_creditors/_competitors)_against_the_company_promoters/director_etc_if_any': '', 'Certification/_Compliances': '', 'If_any_Adverse_Audit_Remarks_in_last_3_years_audited_Annual_Reports': '', 'Position_of_statutory_dues': '', 'Contingent_Liability_if_any': '', 'BO_Remarks_(in_case_of_SIDBI_customer)': '', 'Remarks_of_MSME_Equity_Cell:': '', '(a)_Credibility_of_Issuer': '', '(b)_Credibility_of_Merchant_Banker': '', '(c)_Liability_on_SIDBI': '', '(d)_Proposed_Investment': ''}}]
# IM_TEMPLATE_SCHEMA = [{'placeholders': {'Name_of_Issuer_Company': '', 'MSME_status/_Udyam_Status': '', 'Date_of_Incorporation': '', 'Commencement_of_Operations': '', 'Address_of_Registered_Office': '', 'Address_of_Operating_facilities': ''}}, {'Brief_Background_of_company_and_promoters': ''}, {'Details_about_DRHP_filing_date_approval_date': ''}, {'Details_about_Exchange_on_which_company_is_going_to_be_listed_Merchant_banker_and_Registrar.': ''}, {'placeholders': {'IPO_Date': '', 'Listing_Date': '', 'Face_Value': '', 'Issue_Price_Band': '', 'Lot_Size': '', 'Sale_Type': '', 'Issue_size': '', 'Issue_Type': '', 'Listing_At': '', 'Share_Holding_Pre_Issue': '', 'Share_Holding_Post_Issue': ''}}, {'placeholders': {'QIB_Shares_Offered_Shares_Offered': '', 'Retail_Shares_Offered_Shares_Offered': '', 'NII_Shares_Offered_Shares_Offered': ''}}, {'placeholders': {'Assets_31_March_2025': '', 'Assets_31_Mar_2024': '', 'Assets_31_Mar_2023': '', 'Total_Income_31_March_2025': '', 'Total_Income_31_Mar_2024': '', 'Total_Income_31_Mar_2023': '', 'Profit_After_Tax_31_March_2025': '', 'Profit_After_Tax_31_Mar_2024': '', 'Profit_After_Tax_31_Mar_2023': '', 'Net_Worth_31_March_2025': '', 'Net_Worth_31_Mar_2024': '', 'Net_Worth_31_Mar_2023': '', 'Reserves_and_Surplus_31_March_2025': '', 'Reserves_and_Surplus_31_Mar_2024': '', 'Reserves_and_Surplus_31_Mar_2023': '', 'Total_Borrowing_31_March_2025': '', 'Total_Borrowing_31_Mar_2024': '', 'Total_Borrowing_31_Mar_2023': ''}}, {'Brief_about_company’s_projections_profitability_etc': ''}, {'placeholders': {'Promoter_group_Shareholding_Pattern_Pre-IPO_Nos': '', 'Promoter_group_Shareholding_Pattern_Pre-IPO_%': '', 'Promoter_group_Shareholding_Pattern_Post_IPO_Nos': '', 'Promoter_group_Shareholding_Pattern_Post_IPO_%': '', 'Subtotal_(A)_Shareholding_Pattern_Pre-IPO_Nos': '', 'Subtotal_(A)_Shareholding_Pattern_Pre-IPO_%': '', 'Subtotal_(A)_Shareholding_Pattern_Post_IPO_Nos': '', 'Subtotal_(A)_Shareholding_Pattern_Post_IPO_%': '', 'AIF/Individuals/corporates_Shareholding_Pattern_Pre-IPO_Nos': '', 'AIF/Individuals/corporates_Shareholding_Pattern_Pre-IPO_%': '', 'AIF/Individuals/corporates_Shareholding_Pattern_Post_IPO_Nos': '', 'AIF/Individuals/corporates_Shareholding_Pattern_Post_IPO_%': '', 'IPO_shareholders_Shareholding_Pattern_Pre-IPO_Nos': '', 'IPO_shareholders_Shareholding_Pattern_Pre-IPO_%': '', 'IPO_shareholders_Shareholding_Pattern_Post_IPO_Nos': '', 'IPO_shareholders_Shareholding_Pattern_Post_IPO_%': '', 'Subtotal_(B)_Shareholding_Pattern_Pre-IPO_Nos': '', 'Subtotal_(B)_Shareholding_Pattern_Pre-IPO_%': '', 'Subtotal_(B)_Shareholding_Pattern_Post_IPO_Nos': '', 'Subtotal_(B)_Shareholding_Pattern_Post_IPO_%': '', 'Total_(A+B)_Shareholding_Pattern_Pre-IPO_Nos': '', 'Total_(A+B)_Shareholding_Pattern_Pre-IPO_%': '', 'Total_(A+B)_Shareholding_Pattern_Post_IPO_Nos': '', 'Total_(A+B)_Shareholding_Pattern_Post_IPO_Percentage': ''}}, {'placeholders': {'DRHP_should_have_been_approved_by_any_Stock_Exchange_/_SEBI._Compliance': '', 'Should_have_availed_at_least_one_Debt_products_from_Banks/FIs_and_servicing_it_for_at_least_3_years_prior_to_the_date_of_filing_of_DRHP._and_/_or_should_have_availed_funding_from_any_SEBI_registered_AIFs_at_least_1_year_prior_to_the_date_of_filing_of_DRHP._Compliance': '', 'Net_worth_>_₹_20_crore_Compliance': '', 'EBIDTA_margin_-_₹2.00_cr._or_>_5%_(whichever_is_more)_for_at_least_two_out_of_the_three_most_recent_financial_years_Compliance': '', 'Total_Income_>_₹50_crore_Compliance': '', 'Net_Tangible_Assets_>_₹_2.50_crore_Compliance': '', 'Profitability_(total_PBT_of_3_years)_>_₹_10_crore_Compliance': '', 'Exposure_cap_per_MSME:_Minimum:_₹_1_crore_Maximum:_Up_to_₹_20_crore._Note:_Anchor_investment_shall_not_exceed_50%_of_the_Anchor_portion_of_the_Issue_or_10%_of_the_post_issue_paid-up_capital_of_the_issuer_company_whichever_is_lower._Compliance': '', 'Minimum_Promoters_Shareholding_(Post_Issue)-_60%_In_case_where_the_Company_has_received_any_investment_from_a_SEBI_registered_AIF_the_Investment_Committee_may_take_a_suitable_view_on_the_minimum_level_of_Promoter_Shareholding_below_60%_since_the_AIF_would_want_to_partially_/_fully_exit_in_the_proposed_IPO_/_Offer_for_Sale_(OFS)._Compliance': '', 'Complied_with_Minimum_₹_2_Cr._Bid_criteria_of_SEBI_Compliance': '', 'Complied_with_RBI’s_guideline_limiting_post-issue_shareholding_exposure_to_a_maximum_of_10%._Compliance': ''}}, {'Brief_about_industry': ''}, {'placeholders': {'1_Items': '', '1_Amount_(₹_crore)': '', '1_%': '', '2_Items': '', '2_Amount_(₹_crore)': '', '2_%': '', '3_Items': '', '3_Amount_(₹_crore)': '', '3_%': '', '4_Items': '', '4_Amount_(₹_crore)': '', '4_%': '', 'Total_Items': '', 'Total_Amount_(₹_crore)': '', 'Total_%': ''}}, {'Brief_about_basis_of_valuation': ''}, {'placeholders': {'PAT_(in_lakh)_Mar-25': '', 'PAT_(in_lakh)_Mar-24': '', 'PAT_(in_lakh)_Mar-23': '', 'EPS_Mar-25': '', 'EPS_Mar-24': '', 'EPS_Mar-23': '', 'EPS_after_adjusting_Bonus_Mar-25': '', 'EPS_after_adjusting_Bonus_Mar-24': '', 'EPS_after_adjusting_Bonus_Mar-23': '', 'EPS_after_Proposed_Issue_Mar-25': '', 'EPS_after_Proposed_Issue_Mar-24': '', 'EPS_after_Proposed_Issue_Mar-23': '', 'Weight_Mar-25': '', 'Weight_Mar-24': '', 'Weight_Mar-23': '', 'EPS_Product_Mar-25': '', 'EPS_Product_Mar-24': '', 'EPS_Product_Mar-23': '', 'PE_Ratio_(2025_EPS)_Mar-25': '', 'PE_Ratio_(2025_EPS)_Mar-24': '', 'PE_Ratio_(2025_EPS)_Mar-23': '', 'PE_Ratio_Weighted_Average_EPS_Mar-25': '', 'PE_Ratio_Weighted_Average_EPS_Mar-24': '', 'PE_Ratio_Weighted_Average_EPS_Mar-23': '', 'PE_Ratio_Post_issue_based_on_FY_2025_Earning_Mar-25': '', 'PE_Ratio_Post_issue_based_on_FY_2025_Earning_Mar-24': '', 'PE_Ratio_Post_issue_based_on_FY_2025_Earning_Mar-23': '', 'Price_(Rs)_Mar-25': '', 'Price_(Rs)_Mar-24': '', 'Price_(Rs)_Mar-23': '', 'Issue_Size_(Rs_Cr)_-_Public_Mar-25': '', 'Issue_Size_(Rs_Cr)_-_Public_Mar-24': '', 'Issue_Size_(Rs_Cr)_-_Public_Mar-23': '', 'Post_Issue_Valuation_of_Mar-25': '', 'Post_Issue_Valuation_of_Mar-24': '', 'Post_Issue_Valuation_of_Mar-23': '', 'Pre_Issue_PBV_Mar-25': '', 'Pre_Issue_PBV_Mar-24': '', 'Pre_Issue_PBV_Mar-23': '', 'Post_Issue_Book_Value_Mar-25': '', 'Post_Issue_Book_Value_Mar-24': '', 'Post_Issue_Book_Value_Mar-23': '', 'Post_Issue_PBV_Mar-25': '', 'Post_Issue_PBV_Mar-24': '', 'Post_Issue_PBV_Mar-23': '', 'Post_Issue_Networth_(cr)_Mar-25': '', 'Post_Issue_Networth_(cr)_Mar-24': '', 'Post_Issue_Networth_(cr)_Mar-23': ''}}, {'placeholders': {'Company_Name': '', 'FV': '', 'Net_worth': '', 'Sales': '', 'PAT': '', 'EPS': '', 'Issue_Price': '', 'P/E': '', 'B.V_Pre-Issue(`)': '', 'PBV': '', 'RONW_(%)': '', '(`)': '', '(x)': ''}},{'total_issue_size_price_in_crore':''},{'Number_of_shares': ''}, {'Number_of_lots': ''}, {'Number_of_shares': ''}, {'about_promoters_experience': ''}, {'about_consistent_in_profitability': ''}, {'about_entry_barriers': ''}, {'any_other_rationale': ''}, {'About_Lead_and_Co-lead_merchant_bankers_their_past_track_record': ''}, {'Brief_about_Market_Maker_and_their_track_record': ''}, {'On_the_basis_of_Credibility_of_issuer_Mechant_banker_potential_of_returns_Good_investors_backing_etc': ''}, {'Date': ''},{'placeholders': {'Compliance_of_KYC_guidelines': '', 'KYC_Risk_Score': '', 'Verification_of_defaulters’_lists_Watchout_Investor_website_Income_Tax_defaulter_Search_etc.': '', 'Verification_of_CIBIL_reports_For_Company_and_Promoters': '', 'JOCATA_Score': '', 'Market_Inquiries/_Feedback_from_references': '', 'Connected_Lending': '', 'Pending_court_cases_initiated_by_other_banks/_FIs_against_the_company_promoters/director_etc_if_any': '', 'Pending_court_cases_initiated_by_third_parties_(Debtors/_creditors/_competitors)_against_the_company_promoters/director_etc_if_any': '', 'Certification/_Compliances': '', 'If_any_Adverse_Audit_Remarks_in_last_3_years_audited_Annual_Reports': '', 'Position_of_statutory_dues': '', 'Contingent_Liability_if_any': '', 'BO_Remarks_(in_case_of_SIDBI_customer)': '', 'Remarks_of_MSME_Equity_Cell:': '', '(a)_Credibility_of_Issuer': '', '(b)_Credibility_of_Merchant_Banker': '', '(c)_Liability_on_SIDBI': '', '(d)_Proposed_Investment': ''}}]
IM_TEMPLATE_SCHEMA = [{'placeholders': {'Name_of_Issuer_Company': '', 'MSME_status/_Udyam_Status': '', 'Date_of_Incorporation': '', 'Commencement_of_Operations': '', 'Address_of_Registered_Office': '', 'Address_of_Operating_facilities': ''}}, {'Brief_Background_of_company_and_promoters': ''}, {'Details_about_DRHP_filing_date_approval_date': ''}, {'Details_about_Exchange_on_which_company_is_going_to_be_listed_Merchant_banker_and_Registrar.': ''}, {'placeholders': {'IPO_Date': '', 'Listing_Date': '', 'Face_Value': '', 'Issue_Price_Band': '', 'Lot_Size': '', 'Sale_Type': '', 'Issue_size': '', 'Issue_Type': '', 'Listing_At': '', 'Share_Holding_Pre_Issue': '', 'Share_Holding_Post_Issue': ''}}, {'placeholders': {'QIB_Shares_Offered_Shares_Offered': '', 'Retail_Shares_Offered_Shares_Offered': '', 'NII_Shares_Offered_Shares_Offered': ''}}, {'placeholders': {'Assets_31_March_2025': '', 'Assets_31_Mar_2024': '', 'Assets_31_Mar_2023': '', 'Total_Income_31_March_2025': '', 'Total_Income_31_Mar_2024': '', 'Total_Income_31_Mar_2023': '', 'Profit_After_Tax_31_March_2025': '', 'Profit_After_Tax_31_Mar_2024': '', 'Profit_After_Tax_31_Mar_2023': '', 'Net_Worth_31_March_2025': '', 'Net_Worth_31_Mar_2024': '', 'Net_Worth_31_Mar_2023': '', 'Reserves_and_Surplus_31_March_2025': '', 'Reserves_and_Surplus_31_Mar_2024': '', 'Reserves_and_Surplus_31_Mar_2023': '', 'Total_Borrowing_31_March_2025': '', 'Total_Borrowing_31_Mar_2024': '', 'Total_Borrowing_31_Mar_2023': ''}}, {'Brief_about_company’s_projections_profitability_etc': ''}, {'placeholders': {'Promoter_Shareholding_Pattern_Pre-IPO_Number_of_Shares': '', 'Promoter_Shareholding_Pattern_Pre-IPO_Percentage': '', 'Promoter_Shareholding_Pattern_Post_IPO_Number_of_Shares': '', 'Promoter_Shareholding_Pattern_Post_IPO_Percentage': '','Promoter_group_Shareholding_Pattern_Pre-IPO_Number_of_Shares': '', 'Promoter_group_Shareholding_Pattern_Pre-IPO_Percentage': '', 'Promoter_group_Shareholding_Pattern_Post_IPO_Number_of_Shares': '', 'Promoter_group_Shareholding_Pattern_Post_IPO_Percentage': '', 'Subtotal_(A)_Shareholding_Pattern_Pre-IPO_Number_of_Shares': '', 'Subtotal_(A)_Shareholding_Pattern_Pre-IPO_Percentage': '', 'Subtotal_(A)_Shareholding_Pattern_Post_IPO_Number_of_Shares': '', 'Subtotal_(A)_Shareholding_Pattern_Post_IPO_Percentage': '', 'AIF/Individuals/corporates_Shareholding_Pattern_Pre-IPO_Number_of_Shares': '', 'AIF/Individuals/corporates_Shareholding_Pattern_Pre-IPO_Percentage': '', 'AIF/Individuals/corporates_Shareholding_Pattern_Post_IPO_Number_of_Shares': '', 'AIF/Individuals/corporates_Shareholding_Pattern_Post_IPO_Percentage': '', 'IPO_shareholders_Shareholding_Pattern_Pre-IPO_Number_of_Shares': '', 'IPO_shareholders_Shareholding_Pattern_Pre-IPO_Percentage': '', 'IPO_shareholders_Shareholding_Pattern_Post_IPO_Number_of_Shares': '', 'IPO_shareholders_Shareholding_Pattern_Post_IPO_Percentage': '', 'Subtotal_(B)_Shareholding_Pattern_Pre-IPO_Number_of_Shares': '', 'Subtotal_(B)_Shareholding_Pattern_Pre-IPO_Percentage': '', 'Subtotal_(B)_Shareholding_Pattern_Post_IPO_Number_of_Shares': '', 'Subtotal_(B)_Shareholding_Pattern_Post_IPO_Percentage': '', 'Total_(A+B)_Shareholding_Pattern_Pre-IPO_Number_of_Shares': '', 'Total_(A+B)_Shareholding_Pattern_Pre-IPO_Percentage': '', 'Total_(A+B)_Shareholding_Pattern_Post_IPO_Number_of_Shares': '', 'Total_(A+B)_Shareholding_Pattern_Post_IPO_Percentage': ''}}, {'placeholders': {'DRHP_should_have_been_approved_by_any_Stock_Exchange_/_SEBI._Compliance': '', 'Should_have_availed_at_least_one_Debt_products_from_Banks/FIs_and_servicing_it_for_at_least_3_years_prior_to_the_date_of_filing_of_DRHP._and_/_or_should_have_availed_funding_from_any_SEBI_registered_AIFs_at_least_1_year_prior_to_the_date_of_filing_of_DRHP._Compliance': '', 'Net_worth_>_₹_20_crore_Compliance': '', 'EBIDTA_margin_-_₹2.00_cr._or_>_5%_(whichever_is_more)_for_at_least_two_out_of_the_three_most_recent_financial_years_Compliance': '', 'Total_Income_>_₹50_crore_Compliance': '', 'Net_Tangible_Assets_>_₹_2.50_crore_Compliance': '', 'Profitability_(total_PBT_of_3_years)_>_₹_10_crore_Compliance': '', 'Minimum_Promoters_Shareholding_(Post_Issue)-_60%_In_case_where_the_Company_has_received_any_investment_from_a_SEBI_registered_AIF_the_Investment_Committee_may_take_a_suitable_view_on_the_minimum_level_of_Promoter_Shareholding_below_60%_since_the_AIF_would_want_to_partially_/_fully_exit_in_the_proposed_IPO_/_Offer_for_Sale_(OFS)._Compliance': ''}}, {'Brief_about_industry': ''}, {'placeholders': {'1_Items': '', '1_Amount_(₹_crore)': '', '1_%': '', '2_Items': '', '2_Amount_(₹_crore)': '', '2_%': '', '3_Items': '', '3_Amount_(₹_crore)': '', '3_%': '', '4_Items': '', '4_Amount_(₹_crore)': '', '4_%': '', 'Total_Items': '', 'Total_Amount_(₹_crore)': '', 'Total_%': ''}}, {'Brief_about_basis_of_valuation': ''}, {'placeholders': {'PAT_(`_in_lakh)_Mar-25': '', 'PAT_(`_in_lakh)_Mar-24': '', 'PAT_(`_in_lakh)_Mar-23': '', 'EPS_Mar-25': '', 'EPS_Mar-24': '', 'EPS_Mar-23': '', 'EPS_after_adjusting_Bonus_Mar-25': '', 'EPS_after_adjusting_Bonus_Mar-24': '', 'EPS_after_adjusting_Bonus_Mar-23': '', 'EPS_after_Proposed_Issue_Mar-25': '', 'EPS_after_Proposed_Issue_Mar-24': '', 'EPS_after_Proposed_Issue_Mar-23': '', 'Weight_Mar-25': '', 'Weight_Mar-24': '', 'Weight_Mar-23': '', 'EPS_Product_Mar-25': '', 'EPS_Product_Mar-24': '', 'EPS_Product_Mar-23': '','weighted_avg_eps': '', 'PE_Ratio_(2025_EPS)_Mar-25': '', 'PE_Ratio_(2025_EPS)_Mar-24': '', 'PE_Ratio_(2025_EPS)_Mar-23': '', 'PE_Ratio_Weighted_Average_EPS_Mar-25': '', 'PE_Ratio_Weighted_Average_EPS_Mar-24': '', 'PE_Ratio_Weighted_Average_EPS_Mar-23': '', 'PE_Ratio_Post_issue_based_on_FY_2025_Earning_Mar-25': '', 'PE_Ratio_Post_issue_based_on_FY_2025_Earning_Mar-24': '', 'PE_Ratio_Post_issue_based_on_FY_2025_Earning_Mar-23': '', 'Price_(Rs)_Mar-25': '', 'Price_(Rs)_Mar-24': '', 'Price_(Rs)_Mar-23': '', 'Issue_Size_(Rs_Cr)_-_Public_Mar-25': '', 'Issue_Size_(Rs_Cr)_-_Public_Mar-24': '', 'Issue_Size_(Rs_Cr)_-_Public_Mar-23': '', 'Post_Issue_Valuation_of_Company_(Rs_Cr)_Mar-25': '', 'Post_Issue_Valuation_of_Company_(Rs_Cr)_Mar-24': '', 'Post_Issue_Valuation_of_Company_(Rs_Cr)_Mar-23': '', 'Pre_Issue_PBV_Mar-25': '', 'Pre_Issue_PBV_Mar-24': '', 'Pre_Issue_PBV_Mar-23': '', 'Post_Issue_Book_Value_Mar-25': '', 'Post_Issue_Book_Value_Mar-24': '', 'Post_Issue_Book_Value_Mar-23': '', 'Post_Issue_PBV_Mar-25': '', 'Post_Issue_PBV_Mar-24': '', 'Post_Issue_PBV_Mar-23': '', 'Post_Issue_Networth_(`_cr)_Mar-25': '', 'Post_Issue_Networth_(`_cr)_Mar-24': '', 'Post_Issue_Networth_(`_cr)_Mar-23': ''}}, {'placeholders': {'1._Company_Name': '', '1._FV': '', '1._Net_worth': '', '1._Sales': '', '1._PAT': '', '1._EPS': '', '1._Issue_Price': '', '1._P/E': '', '1._B.V_Pre-Issue': '', '1._PBV': '', '1._RONW(%)': '','Peer_1._Company_Name': '', 'Peer_1._FV': '', 'Peer_1._Net_worth': '', 'Peer_1._Sales': '', 'Peer_1._PAT': '', 'Peer_1._EPS': '', 'Peer_1._Issue_Price': '', 'Peer_1._P/E': '', 'Peer_1._B.V_Pre-Issue': '', 'Peer_1._PBV': '', 'Peer_1._RONW(%)': '','Peer_2._Company_Name': '', 'Peer_2._FV': '', 'Peer_2._Net_worth': '', 'Peer_2._Sales': '', 'Peer_2._PAT': '', 'Peer_2._EPS': '', 'Peer_2._Issue_Price': '', 'Peer_2._P/E': '', 'Peer_2._B.V_Pre-Issue': '', 'Peer_2._PBV': '', 'Peer_2._RONW(%)': ''}},{'total_issue_size_price':''},{'Number_of_shares': ''}, {'about_promoters_experience': ''}, {'about_consistent_in_profitability': ''}, {'about_entry_barriers': ''}, {'any_other_rationale': ''}, {'About_Lead_and_Co-lead_merchant_bankers_their_past_track_record': ''}, {'Brief_about_Market_Maker_and_their_track_record': ''}, {'On_the_basis_of_Credibility_of_issuer_Mechant_banker_potential_of_returns_Good_investors_backing_etc': ''}, {'Date': ''},{'Offer_for_sale_size': ''},{'placeholders': {'Compliance_of_KYC_guidelines': '', 'Connected_Lending_From_related_party_transaction_sub_section under_summary_of_issue_document': '', 'Pending_court_cases_initiated_by_other_banks/_FIs_against_the_company_promoters/director_etc_if_any': '', 'Pending_court_cases_initiated_by_third_parties_(Debtors/_creditors/_competitors)_against_the_company_promoters/director_etc_if_any': '', 'Certification/_Compliances': '', 'If_any_Adverse_Audit_Remarks_in_last_3_years_audited_Annual_Reports': '', 'Position_of_statutory_dues': '', 'Contingent_Liability_if_any': '', 'BO_Remarks_(in_case_of_SIDBI_customer)': ''}}]
# IM_TEMPLATE_SCHEMA = [{'placeholders': {'Name_of_Issuer_Company': '', 'MSME_status/_Udyam_Status': '', 'Date_of_Incorporation': '', 'Commencement_of_Operations': '', 'Address_of_Registered_Office': '', 'Address_of_Operating_facilities': ''}}, {'Brief_Background_of_company_and_promoters': ''}, {'Details_about_DRHP_filing_date_approval_date': ''}, {'Details_about_Exchange_on_which_company_is_going_to_be_listed_Merchant_banker_and_Registrar.': ''}, {'placeholders': {'IPO_Date': '', 'Listing_Date': '', 'Face_Value': '', 'Issue_Price_Band': '', 'Lot_Size': '', 'Sale_Type': '', 'Issue_size': '', 'Issue_Type': '', 'Listing_At': '', 'Share_Holding_Pre_Issue': '', 'Share_Holding_Post_Issue': ''}}, {'placeholders': {'QIB_Shares_Offered_Shares_Offered': '', 'Retail_Shares_Offered_Shares_Offered': '', 'NII_Shares_Offered_Shares_Offered': ''}}, {'placeholders': {'Assets_31_March_2025': '', 'Assets_31_Mar_2024': '', 'Assets_31_Mar_2023': '', 'Total_Income_31_March_2025': '', 'Total_Income_31_Mar_2024': '', 'Total_Income_31_Mar_2023': '', 'Profit_After_Tax_31_March_2025': '', 'Profit_After_Tax_31_Mar_2024': '', 'Profit_After_Tax_31_Mar_2023': '', 'Net_Worth_31_March_2025': '', 'Net_Worth_31_Mar_2024': '', 'Net_Worth_31_Mar_2023': '', 'Reserves_and_Surplus_31_March_2025': '', 'Reserves_and_Surplus_31_Mar_2024': '', 'Reserves_and_Surplus_31_Mar_2023': '', 'Total_Borrowing_31_March_2025': '', 'Total_Borrowing_31_Mar_2024': '', 'Total_Borrowing_31_Mar_2023': ''}}, {'Brief_about_company’s_projections_profitability_etc': ''}, {'placeholders': {'Promoter_group_Shareholding_Pattern_Pre-IPO_Nos': '', 'Promoter_group_Shareholding_Pattern_Pre-IPO_%': '', 'Promoter_group_Shareholding_Pattern_Post_IPO_Nos': '', 'Promoter_group_Shareholding_Pattern_Post_IPO_%': '', 'Subtotal_(A)_Shareholding_Pattern_Pre-IPO_Nos': '', 'Subtotal_(A)_Shareholding_Pattern_Pre-IPO_%': '', 'Subtotal_(A)_Shareholding_Pattern_Post_IPO_Nos': '', 'Subtotal_(A)_Shareholding_Pattern_Post_IPO_%': '', 'AIF/Individuals/corporates_Shareholding_Pattern_Pre-IPO_Nos': '', 'AIF/Individuals/corporates_Shareholding_Pattern_Pre-IPO_%': '', 'AIF/Individuals/corporates_Shareholding_Pattern_Post_IPO_Nos': '', 'AIF/Individuals/corporates_Shareholding_Pattern_Post_IPO_%': '', 'IPO_shareholders_Shareholding_Pattern_Pre-IPO_Nos': '', 'IPO_shareholders_Shareholding_Pattern_Pre-IPO_%': '', 'IPO_shareholders_Shareholding_Pattern_Post_IPO_Nos': '', 'IPO_shareholders_Shareholding_Pattern_Post_IPO_%': '', 'Subtotal_(B)_Shareholding_Pattern_Pre-IPO_Nos': '', 'Subtotal_(B)_Shareholding_Pattern_Pre-IPO_%': '', 'Subtotal_(B)_Shareholding_Pattern_Post_IPO_Nos': '', 'Subtotal_(B)_Shareholding_Pattern_Post_IPO_%': '', 'Total_(A+B)_Shareholding_Pattern_Pre-IPO_Nos': '', 'Total_(A+B)_Shareholding_Pattern_Pre-IPO_%': '', 'Total_(A+B)_Shareholding_Pattern_Post_IPO_Nos': '', 'Total_(A+B)_Shareholding_Pattern_Post_IPO_%': ''}}, {'placeholders': {'DRHP_should_have_been_approved_by_any_Stock_Exchange_/_SEBI._Compliance': '', 'Should_have_availed_at_least_one_Debt_products_from_Banks/FIs_and_servicing_it_for_at_least_3_years_prior_to_the_date_of_filing_of_DRHP._and_/_or_should_have_availed_funding_from_any_SEBI_registered_AIFs_at_least_1_year_prior_to_the_date_of_filing_of_DRHP._Compliance': '', 'Net_worth_>_₹_20_crore_Compliance': '', 'EBIDTA_margin_-_₹2.00_cr._or_>_5%_(whichever_is_more)_for_at_least_two_out_of_the_three_most_recent_financial_years_Compliance': '', 'Total_Income_>_₹50_crore_Compliance': '', 'Net_Tangible_Assets_>_₹_2.50_crore_Compliance': '', 'Profitability_(total_PBT_of_3_years)_>_₹_10_crore_Compliance': '', 'Exposure_cap_per_MSME:_Minimum:_₹_1_crore_Maximum:_Up_to_₹_20_crore._Note:_Anchor_investment_shall_not_exceed_50%_of_the_Anchor_portion_of_the_Issue_or_10%_of_the_post_issue_paid-up_capital_of_the_issuer_company_whichever_is_lower._Compliance': '', 'Minimum_Promoters_Shareholding_(Post_Issue)-_60%_In_case_where_the_Company_has_received_any_investment_from_a_SEBI_registered_AIF_the_Investment_Committee_may_take_a_suitable_view_on_the_minimum_level_of_Promoter_Shareholding_below_60%_since_the_AIF_would_want_to_partially_/_fully_exit_in_the_proposed_IPO_/_Offer_for_Sale_(OFS)._Compliance': '', 'Complied_with_Minimum_₹_2_Cr._Bid_criteria_of_SEBI_Compliance': '', 'Complied_with_RBI’s_guideline_limiting_post-issue_shareholding_exposure_to_a_maximum_of_10%._Compliance': ''}}, {'Brief_about_industry': ''}, {'placeholders': {'1_Items': '', '1_Amount_(₹_crore)': '', '1_%': '', '2_Items': '', '2_Amount_(₹_crore)': '', '2_%': '', '3_Items': '', '3_Amount_(₹_crore)': '', '3_%': '', '4_Items': '', '4_Amount_(₹_crore)': '', '4_%': '', 'Total_Items': '', 'Total_Amount_(₹_crore)': '', 'Total_%': ''}}, {'Brief_about_basis_of_valuation': ''}, {'placeholders': {'PAT_(in_lakh)_Mar-25': '', 'PAT_(in_lakh)_Mar-24': '', 'PAT_(in_lakh)_Mar-23': '', 'EPS_Mar-25': '', 'EPS_Mar-24': '', 'EPS_Mar-23': '', 'EPS_after_adjusting_Bonus_Mar-25': '', 'EPS_after_adjusting_Bonus_Mar-24': '', 'EPS_after_adjusting_Bonus_Mar-23': '', 'EPS_after_Proposed_Issue_Mar-25': '', 'EPS_after_Proposed_Issue_Mar-24': '', 'EPS_after_Proposed_Issue_Mar-23': '', 'Weight_Mar-25': '', 'Weight_Mar-24': '', 'Weight_Mar-23': '', 'EPS_Product_Mar-25': '', 'EPS_Product_Mar-24': '', 'EPS_Product_Mar-23': '', 'PE_Ratio_(2025_EPS)_Mar-25': '', 'PE_Ratio_(2025_EPS)_Mar-24': '', 'PE_Ratio_(2025_EPS)_Mar-23': '', 'PE_Ratio_Weighted_Average_EPS_Mar-25': '', 'PE_Ratio_Weighted_Average_EPS_Mar-24': '', 'PE_Ratio_Weighted_Average_EPS_Mar-23': '', 'PE_Ratio_Post_issue_based_on_FY_2025_Earning_Mar-25': '', 'PE_Ratio_Post_issue_based_on_FY_2025_Earning_Mar-24': '', 'PE_Ratio_Post_issue_based_on_FY_2025_Earning_Mar-23': '', 'Price_(Rs)_Mar-25': '', 'Price_(Rs)_Mar-24': '', 'Price_(Rs)_Mar-23': '', 'Issue_Size_(Rs_Cr)_-_Public_Mar-25': '', 'Issue_Size_(Rs_Cr)_-_Public_Mar-24': '', 'Issue_Size_(Rs_Cr)_-_Public_Mar-23': '', 'Post_Issue_Valuation_of_Mar-25': '', 'Post_Issue_Valuation_of_Mar-24': '', 'Post_Issue_Valuation_of_Mar-23': '', 'Company_(Rs_Cr)_Mar-25': '', 'Company_(Rs_Cr)_Mar-24': '', 'Company_(Rs_Cr)_Mar-23': '', 'Pre_Issue_PBV_Mar-25': '', 'Pre_Issue_PBV_Mar-24': '', 'Pre_Issue_PBV_Mar-23': '', 'Post_Issue_Book_Value_Mar-25': '', 'Post_Issue_Book_Value_Mar-24': '', 'Post_Issue_Book_Value_Mar-23': '', 'Post_Issue_PBV_Mar-25': '', 'Post_Issue_PBV_Mar-24': '', 'Post_Issue_PBV_Mar-23': '', 'Post_Issue_Networth_(cr)_Mar-25': '', 'Post_Issue_Networth_(cr)_Mar-24': '', 'Post_Issue_Networth_(cr)_Mar-23': ''}}, {'placeholders': {'Company_Name': '', 'FV': '', 'Net_worth': '', 'Sales': '', 'PAT': '', 'EPS': '', 'Issue_Price': '', 'P/E': '', 'B.V_Pre-Issue(`)': '', 'PBV': '', 'RONW_(%)': '', '(`)': '', '(x)': ''}},{'total_issue_size_price_in_crore':''},{'Number_of_shares': ''}, {'Number_of_lots': ''}, {'Number_of_shares': ''}, {'about_promoters_experience': ''}, {'about_consistent_in_profitability': ''}, {'about_entry_barriers': ''}, {'any_other_rationale': ''}, {'About_Lead_and_Co-lead_merchant_bankers_their_past_track_record': ''}, {'Brief_about_Market_Maker_and_their_track_record': ''}, {'On_the_basis_of_Credibility_of_issuer_Mechant_banker_potential_of_returns_Good_investors_backing_etc': ''}, {'Date': ''},{'placeholders': {'Compliance_of_KYC_guidelines': '', 'KYC_Risk_Score': '', 'Verification_of_defaulters’_lists_Watchout_Investor_website_Income_Tax_defaulter_Search_etc.': '', 'Verification_of_CIBIL_reports_For_Company_and_Promoters': '', 'JOCATA_Score': '', 'Market_Inquiries/_Feedback_from_references': '', 'Connected_Lending': '', 'Pending_court_cases_initiated_by_other_banks/_FIs_against_the_company_promoters/director_etc_if_any': '', 'Pending_court_cases_initiated_by_third_parties_(Debtors/_creditors/_competitors)_against_the_company_promoters/director_etc_if_any': '', 'Certification/_Compliances': '', 'If_any_Adverse_Audit_Remarks_in_last_3_years_audited_Annual_Reports': '', 'Position_of_statutory_dues': '', 'Contingent_Liability_if_any': '', 'BO_Remarks_(in_case_of_SIDBI_customer)': '', 'Remarks_of_MSME_Equity_Cell:': '', '(a)_Credibility_of_Issuer': '', '(b)_Credibility_of_Merchant_Banker': '', '(c)_Liability_on_SIDBI': '', '(d)_Proposed_Investment': ''}}]
DRHP_TEMPLATE_SCHEMA = [{'placeholders': {'Name_of_Issuer_Company': '', 'MSME_status/_Udyam_Status': '', 'Date_of_Incorporation': '', 'Commencement_of_Operations': '', 'Address_of_Registered_Office': '', 'Address_of_Operating_facilities': ''}}]
# {'total_issue_size_price_in_crore':''},
HTML_TEMPLATE = """

<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Appraisal Note for Anchor Investment (SIDBI MSME Equity)</title>
  <style>
    :root{
      --accent:#005da6; /* SIDBI blue-ish */
      --light:#f5f7fb;
      --text:#1b1f23;
      --muted:#6b7280;
      --border:#d1d5db;
    }
    body{font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, "Noto Sans", "Liberation Sans", sans-serif; color:var(--text); margin:0; background:var(--light);}
    .page{max-width:1000px; margin:40px auto; background:#fff; box-shadow:0 4px 24px rgba(0,0,0,.08); border-radius:12px;}
    header{padding:24px 28px; border-bottom:4px solid var(--accent);}
    header h1{margin:0; font-size:1.5rem;}
    header .subtitle{color:var(--muted); margin-top:6px;}
    main{padding:24px 28px 40px;}
    h2{color:var(--accent); margin:24px 0 8px; font-size:1.15rem;}
    h3{color:var(--text); margin:16px 0 6px; font-size:1rem;}
    .kpi{font-weight:600;}
    table{width:100%; border-collapse:collapse; margin:8px 0 20px;}
    th, td{border:1px solid var(--border); padding:8px 10px; vertical-align:top;}
    th{background:#f9fafb; text-align:left;}
    .note{background:#fff7ed; border:1px solid #fed7aa; padding:10px 12px; border-radius:8px;}
    .grid{display:grid; grid-template-columns: 1fr 1fr; gap:16px;}
    .signoff{margin-top:28px; display:grid; grid-template-columns: 1fr 1fr; gap:16px;}
    .small{font-size:.9rem; color:var(--muted);} 
    .center{text-align:center;}
    .nowrap{white-space:nowrap;}
  </style>
</head>
<body>
  <div class="page">
    <header>
      <h1>Annexure (b)</h1>
      <div class="subtitle"><strong>MSME Equity Vertical, SIDBI, Mumbai</strong></div>
      <div class="subtitle">Proposal for Anchor Investment to Screening Committee</div>
      <div class="subtitle kpi">(Proposed Investment – ~ ₹ ____ crore)</div>
    </header>
    <main>

      <h2>Company snapshot</h2>
      <table>
        <tr><th>Name of Issuer Company</th><th></th></tr>
        <tr><td>MSME status / Udyam Status</td><td></td></tr>
        <tr><td>Date of Incorporation</td><td></td></tr>
        <tr><td>Commencement of Operations</td><td></td></tr>
        <tr><td>Address of Registered Office</td><td></td></tr>
        <tr><td>Address of Operating facilities</td><td></td></tr>
      </table>
      <p></p>
      <h2>Background</h2>
      <p>[Brief Background of company and promoters]</p>

      <h2>About the IPO</h2>
      <p>[Details about DRHP filing date, approval date]</p>
      <p>[Details about Exchange on which company is going to be listed, Merchant banker and Registrar]</p>

      <h3>Other Details of IPO (Tentative)</h3>
      <table>
        <tr><th>IPO Date</th><th></th></tr>
        <tr><td>Listing Date</td><td></td></tr>
        <tr><td>Face Value</td><td></td></tr>
        <tr><td>Issue Price Band</td><td></td></tr>
        <tr><td>Lot Size</td><td></td></tr>
        <tr><td>Sale Type</td><td></td></tr>
        <tr><td>Issue size</td><td></td></tr>
        <tr><td>Issue Type</td><td></td></tr>
        <tr><td>Listing At</td><td></td></tr>
        <tr><td>Share Holding Pre Issue</td><td></td></tr>
        <tr><td>Share Holding Post Issue</td><td></td></tr>
      </table>
      <p></p>
      <h3>IPO distribution (Reservation by category)</h3>
      <table>
        <tr><th>Investor Category</th><th>Shares Offered</th></tr>
        <tr><td>QIB Shares Offered</td><td></td></tr>
        <tr><td>Retail Shares Offered</td><td></td></tr>
        <tr><td>NII Shares Offered</td><td></td></tr>
      </table>
      <p></p>
      <h2>Financial Performance</h2>
      <table>
        <tr>
          <th>Period Ended</th>
          <th>31 Mar 2025</th>
          <th>31 Mar 2024</th>
          <th>31 Mar 2023</th>
        </tr>
        <tr><td>Assets</td><td></td><td></td><td></td></tr>
        <tr><td>Total Income</td><td></td><td></td><td></td></tr>
        <tr><td>Profit After Tax</td><td></td><td></td><td></td></tr>
        <tr><td>Net Worth</td><td></td><td></td><td></td></tr>
        <tr><td>Reserves and Surplus</td><td></td><td></td><td></td></tr>
        <tr><td>Total Borrowing</td><td></td><td></td><td></td></tr>
      </table>
      <p></p>
      <p class="small">Brief About Company's Projections, profitability: [Brief about company’s projections, profitability etc]</p>

      <h2>Shareholding Pattern</h2>
      <table>
        <tr><th colspan="5">Shareholding Pattern</th></tr>
        <tr><td colspan="3"><strong>Pre-IPO</strong></td><td colspan="2"><strong>Post IPO</strong></td></tr>
        <tr><td><strong>Equity shares – FV ₹10/- each</strong></td><td><strong>Nos</strong></td><td><strong>%</strong></td><td><strong>Nos</strong></td><td><strong>%</strong></td></tr>
        <tr><td colspan="5"></td></tr>
        <tr><td colspan="5"><strong>Promoters & Promoter groups</strong></td></tr>
        <tr><td>Promoter</td><td></td><td></td><td></td><td></td></tr>
        <tr><td>Promoter group</td><td></td><td></td><td></td><td></td></tr>
        <tr><td><strong>Subtotal (A)</strong></td><td></td><td></td><td></td><td></td></tr>
        <tr><td colspan="5"><strong>Public/Others</strong></td></tr>
        <tr><td>AIF / Individuals / corporates</td><td></td><td></td><td></td><td></td></tr>
        <tr><td>IPO shareholders</td><td></td><td></td><td></td><td></td></tr>
        <tr><td><strong>Subtotal (B)</strong></td><td></td><td></td><td></td><td></td></tr>
        <tr><td></td><td></td><td>%</td><td></td><td>%</td></tr>
        <tr><td><strong>Total (A+B)</strong></td><td></td><td></td><td></td><td></td></tr>
      </table>
    <p></p>
      <h2>Compliance of Eligibility Criteria</h2>
      <table>
        <tr><th>S.No.</th><th>Criteria</th><th>Compliance</th></tr>
        <tr><td>1</td><td>DRHP should have been approved by any Stock Exchange / SEBI.</td><td></td></tr>
        <tr><td>2</td><td>Should have availed at least one Debt products from Banks/FIs and servicing it for at least 3 years prior to the date of filing of DRHP. and / or should have availed funding from any SEBI registered AIFs at least 1 year prior to the date of filing of DRHP.</td><td></td></tr>
        <tr><td>3</td><td>Net worth &gt; ₹ 20 crore</td><td></td></tr>
        <tr><td>4</td><td>EBIDTA margin – ₹2.00 cr. or &gt; 5% (whichever is more) for at least two out of the three most recent financial years</td><td></td></tr>
        <tr><td>5</td><td>Total Income &gt; ₹50 crore</td><td></td></tr>
        <tr><td>6</td><td>Net Tangible Assets &gt; ₹ 2.50 crore</td><td></td></tr>
        <tr><td>7</td><td>Profitability (total PBT of 3 years) &gt; ₹ 10 crore</td><td></td></tr>
        <tr><td>8</td><td>Exposure cap per MSME: Minimum: ₹ 1 crore; Maximum: Up to ₹ 20 crore. Note: Anchor investment shall not exceed 50% of the Anchor portion of the Issue or 10% of the post issue paid-up capital of the issuer company, whichever is lower.</td><td>Business to fill</td></tr>
        <tr><td>9</td><td>Minimum Promoters Shareholding (Post Issue) – 60%. In case of prior AIF investment, committee may allow below 60% due to AIF exit in IPO/OFS.</td><td></td></tr>
        <tr><td>10</td><td>Complied with Minimum ₹ 2 Cr. Bid criteria of SEBI</td><td>Business to fill</td></tr>
        <tr><td>11</td><td>Complied with RBI guideline limiting post-issue shareholding exposure to a maximum of 10%.</td><td>Business to fill</td></tr>
      </table>
        <p></p>
      <h2>Industry</h2>
      <p>[Brief about industry]</p>

      <h2>Objective of the Issue</h2>
      <p>The company is proposing to use the IPO proceeds as under:</p>
      <table>
        <tr><th>Sr. No.</th><th>Items</th><th>Amount (₹ crore)</th><th>%</th></tr>
        <tr><td>1</td><td></td><td></td><td></td></tr>
        <tr><td>2</td><td></td><td></td><td></td></tr>
        <tr><td>3</td><td></td><td></td><td></td></tr>
        <tr><td>4</td><td></td><td></td><td></td></tr>
        <tr><td>5</td><td></td><td></td><td></td></tr>
        <tr><td>6</td><td></td><td></td><td></td></tr>
        <tr><td>7</td><td></td><td></td><td></td></tr>
        <tr><td></td><td><strong>Total</strong></td><td></td><td>100.00%</td></tr>
      </table>
    <p></p>
      <h2>Basis of the Valuation</h2>
      <p>[Brief about basis of valuation]</p>
      <table>
        <tr><th>Particulars</th><th>Mar-25</th><th>Mar-24</th><th>Mar-23</th></tr>
        <tr><td>PAT (₹ in lakh)</td><td></td><td></td><td></td></tr>
        <tr><td>EPS</td><td></td><td></td><td></td></tr>
        <tr><td>EPS after adjusting Bonus</td><td></td><td></td><td></td></tr>
        <tr><td>EPS after Proposed Issue</td><td></td><td></td><td></td></tr>
        <tr><td>Weight</td><td></td><td></td><td></td></tr>
        <tr><td>EPS Product</td><td></td><td></td><td></td></tr>
        <tr><td>Wtd Avg EPS</td><td colspan="3"><strong></strong></td></tr>
        <tr><td colspan="4"><strong>Sensitivity Analysis (Issue Price &amp; Size)</strong></td></tr>
        <tr><td><strong>PE Ratio (2025 EPS)</strong></td><td></td><td></td><td></td></tr>
        <tr><td>PE Ratio Weighted Average EPS</td><td></td><td></td><td></td></tr>
        <tr><td>PE Ratio Post issue based on FY 2025 Earning</td><td></td><td></td><td></td></tr>
        <tr><td>Price (₹)</td><td></td><td></td><td></td></tr>
        <tr><td>Issue Size (₹ Cr) – Public</td><td></td><td></td><td></td></tr>
        <tr><td>Post Issue Valuation of Company</td><td></td><td></td><td></td></tr>
        <tr><td></td><td colspan="3"></td></tr>
        <tr><td></td><td></td><td></td><td></td></tr>
        <tr><td>Pre Issue PBV</td><td></td><td></td><td></td></tr>
        <tr><td>Post Issue Book Value</td><td></td><td></td><td></td></tr>
        <tr><td>Post Issue PBV</td><td></td><td></td><td></td></tr>
        <tr><td></td><td></td><td></td><td></td></tr>
        <tr><td>Post Issue Networth (₹ cr)</td><td></td><td></td><td></td></tr>
      </table>
        <p></p>
      <h2>Peer comparison</h2>
      <table>
        <tr>
          <th >Sr. No.</th>
          <th>Company Name</th>
          <th>FV</th>
          <th>Net worth</th>
          <th>Sales</th>
          <th>PAT</th>
          <th>EPS</th>
          <th>Issue Price</th>
          <th>P/E</th>
          <th>B.V Pre-Issue (₹)</th>
          <th>PBV</th>
          <th>RONW (%)</th>
        </tr>
        <tr><td>1</td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td></tr>
        <tr><td>2</td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td></tr>
        <tr><td>3</td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td></tr>
      </table>
      <p></p>
      <div class="note">
        <strong>Note: Peer Selection Guideline for Comparative Analysis:</strong>
        In cases where exact peers—those operating in all the same business segments as the issuer—are not available, select peers that match the issuer’s business and segments to the maximum extent possible. Additionally, include another peer that fulfils the remaining criteria. Use a weighted average of these selected peers to compare the issuer company’s financials and other relevant metrics.
      </div>

      <h2>Quantum of Anchor Investment</h2>
      <p>Total Issue Size – ₹ [total issue size price] crore ([Number of shares] shares)<br>
         Anchor Portion – ₹ _____ crore ( _________ shares )<br>
         As per guidelines, minimum application for Anchor Investor in SME IPO is ₹ 2 crore. Proposed application: ____ lot ( ____ shares) amounting to ₹ _____/-.
      </p>

      <h2>Rationale for Investment</h2>
      <ul>
        <li>Promoters Experience: [about promoters experience]</li>
        <li>Consistency in Profitability: [about consistent in profitability]</li>
        <li>Entry Barriers: [about entry barriers]</li>
        <li>Any other rationale: [any other rationale]</li>
      </ul>

      <h2>Merchant Banker – Track Record</h2>
      <p>[About Lead and Co-lead merchant bankers their past track record]</p>
 

      <h2>Market Maker – Comments on Track Record</h2>
      <p>[Brief about Market Maker and their track record]</p>

      <h2>Recommendation</h2>
      <p>[Based on credibility of issuer, merchant banker, potential returns, investor backing etc.]<br>
      In view of the above, the proposal of Anchor Investment of upto ₹ _____/- is recommended to Screening Committee–EI for approval. Post recommendation, the proposal will be placed before the Investment Committee.</p>

      <p>Submitted for recommendation, please.</p>

      <div class="signoff">
        <div>
          <strong>AM/Mgr (Name)</strong><br>
          <span class="small">[Date]</span>
        </div>
        <div>
          <strong>AGM/DGM (Name)</strong>
        </div>
      </div>
      <p><strong>GM/CGM, MEC (Name) Pl.</strong><br><strong>Screening committee–EI</strong></p>

      <h2>Due-Diligence of the company and promoters</h2>
      <h3>KYC &amp; Reference check / verification details</h3>
      <table>
        <tr><th>1</th><th>Compliance of KYC guidelines</th><th></th></tr>
        <tr><td>2</td><td>KYC Risk Score</td><td>Business to fill</td></tr>
        <tr><td>3</td><td>Verification of defaulters’ lists, Watchout Investor website, Income Tax defaulter search, etc.</td><td>Business to fill</td></tr>
        <tr><td>4</td><td>Verification of CIBIL reports – Company and Promoters</td><td>Business to fill</td></tr>
        <tr><td>5</td><td>JOCATA Score</td><td>Business to fill</td></tr>
        <tr><td>6</td><td>Market inquiries / Feedback from references</td><td>Business to fill</td></tr>
        <tr><td>7</td><td>Connected Lending</td><td></td></tr>
        <tr><td>8</td><td>Pending court cases initiated by other banks/FIs against the company, promoters/director etc.</td><td></td></tr>
        <tr><td>9</td><td>Pending court cases initiated by third parties (Debtors/creditors/competitors)</td><td></td></tr>
        <tr><td>10</td><td>Certification/Compliances</td><td></td></tr>
        <tr><td>11</td><td>If any Adverse Audit Remarks in last 3 years audited Annual Reports</td><td></td></tr>
        <tr><td>12</td><td>Position of statutory dues</td><td></td></tr>
        <tr><td>13</td><td>Contingent Liability if any</td><td></td></tr>
        <tr><td>14</td><td>BO Remarks (in case of SIDBI customer)</td><td></td></tr>
        <tr><td>15</td><td>Remarks of MSME Equity Cell:</td><td>Business to fill</td></tr>
        <tr><td></td><td>(a)Credibility of Issuer</td><td>Business to fill</td></tr>
        <tr><td></td><td>(b)Credibility of Merchant Banker</td><td>Business to fill</td></tr>
        <tr><td></td><td>(c)Liability on SIDBI</td><td>Business to fill</td></tr>
        <tr><td></td><td>(d)Proposed Investment</td><td>Business to fill</td></tr>
      </table>

    </main>
  </div>
</body>
</html>


"""
# ================== SESSION STATE INIT ==================
for key, default in [
    ("template_json", None),
    ("pdf_text", None),
    ("docx_uploaded_name", None),
    ("pdf_uploaded_name", None),
    ("filled_html", None),
]:
    if key not in st.session_state:
        st.session_state[key] = default

# ================== HELPERS ==================
def strip_tags(html: str) -> str:
    """Very simple HTML tag stripper."""
    return re.sub(r"<[^>]+>", "", html)



# def html_to_docx_bytes(html: str) -> bytes:
    """
    Convert HTML (like your IM template) to a reasonably structured DOCX.
    - h1/h2/h3 -> Word headings
    - p       -> paragraphs
    - ul/ol   -> bullet/numbered lists
    - table   -> grid table with header row
    """
    soup = BeautifulSoup(html, "html.parser")
    document = Document()

    # Optional: add main title from <title> if present
    if soup.title and soup.title.string:
        title_para = document.add_paragraph(soup.title.string.strip())
        title_para.style = "Title"
        title_para.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Work on <body> if present, else whole soup
    root: Tag = soup.body if soup.body else soup

    def handle_element(el: Tag, doc: Document):
        # Skip non-tag nodes handled via parent
        if isinstance(el, NavigableString):
            text = el.strip()
            if text:
                doc.add_paragraph(text)
            return

        # Headings
        if el.name in ["h1", "h2", "h3"]:
            text = el.get_text(strip=True)
            if not text:
                return
            level = {"h1": 1, "h2": 2, "h3": 3}[el.name]
            doc.add_heading(text, level=level)
            return

        # Paragraphs
        if el.name == "p":
            text = el.get_text(" ", strip=True)
            if text:
                doc.add_paragraph(text)
            return

        # Unordered list (bullets)
        if el.name == "ul":
            for li in el.find_all("li", recursive=False):
                text = li.get_text(" ", strip=True)
                if text:
                    p = doc.add_paragraph(text)
                    p.style = "List Bullet"
            return

        # Ordered list (numbers)
        if el.name == "ol":
            for li in el.find_all("li", recursive=False):
                text = li.get_text(" ", strip=True)
                if text:
                    p = doc.add_paragraph(text)
                    p.style = "List Number"
            return

        # Tables
        if el.name == "table":
            rows = el.find_all("tr")
            if not rows:
                return

            # Determine max columns from first row with cells
            max_cols = 0
            row_cells = []
            for r in rows:
                cells = r.find_all(["th", "td"], recursive=False)
                if cells:
                    row_cells.append(cells)
                    max_cols = max(max_cols, len(cells))

            if max_cols == 0:
                return

            table = doc.add_table(rows=len(row_cells), cols=max_cols)
            table.style = "Table Grid"

            for r_idx, cells in enumerate(row_cells):
                row = table.rows[r_idx]
                for c_idx in range(max_cols):
                    cell = row.cells[c_idx]
                    if c_idx < len(cells):
                        text = cells[c_idx].get_text(" ", strip=True)
                        cell.text = text

            return

        # Structural wrappers: handle children
        if el.name in ["div", "section", "main", "header", "footer"]:
            for child in el.children:
                if isinstance(child, (Tag, NavigableString)):
                    handle_element(child, doc)
            return

        # Fallback: just process children
        for child in el.children:
            if isinstance(child, (Tag, NavigableString)):
                handle_element(child, doc)

    # Process top-level children of root
    for child in root.children:
        if isinstance(child, (Tag, NavigableString)):
            handle_element(child, document)

    # Save to bytes
    buf = BytesIO()
    document.save(buf)
    buf.seek(0)
    return buf.getvalue()
   

left,right = st.columns([0.8,2.2])

with left:
    # st.header("Template & Inputs")
    template_choice = st.selectbox(
        "Select Template",
        ["IM Template","DRHP Template","Custom(DOCX Upload)"],
        index=0
    )

    # Set template_json form dropdown
    if template_choice =="IM Template":
        st.session_state["template_json"]=IM_TEMPLATE_SCHEMA
        # st.info("Using built-in IM Template JSON Schema")
        enable_docx_upload = False
    elif template_choice == "DRHP Template":
        st.session_state["template_json"]=DRHP_TEMPLATE_SCHEMA
        # st.info("Using built-in DRHP Template JSON Schema")
        enable_docx_upload = False
    else:
        enable_docx_upload=True
        st.session_state["template_json"]=None

    # DOCX upload only if custom selected
    if enable_docx_upload:
        docx_file = st.file_uploader("Upload IM DOCX Template file", type=["docx"], key="docx_file")
        if docx_file is not None:
            if st.session_state["docx_uploaded_name"] != docx_file.name:
                try:
                    files = {
                        "docx_file": (
                            docx_file.name,
                            docx_file.read(),
                            "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        )
                    }
                    resp = requests.post(DOCX_BACKEND_URL, files=files, timeout=900)
                    resp.raise_for_status()
                    data = resp.json()
                    print(data)
                    # Expect either {"template_json": {...}} or directly {...}
                    if isinstance(data, dict) and "template_json" in data:
                        st.session_state["template_json"] = data["template_json"]
                    else:   
                        st.session_state["template_json"] = data
                    st.session_state["docx_uploaded_name"] = docx_file.name
                    st.success("Template DOCX uploaded successfully")
                except Exception as e:
                    st.error(f"Error calling docx-to-json backend: {e}")
        else:
            st.session_state["docx_uploaded_name"] = None
            
    else:
        st.session_state["docx_uploaded_name"] = None



    # st.subheader("Step 2: Upload PDF document")
    pdf_file = st.file_uploader("Upload a PDF file", type=["pdf"], key="pdf_file")

    if pdf_file is not None:
        if st.session_state["pdf_uploaded_name"] != pdf_file.name:
            try:
                files = {
                    "pdf_file": (
                        pdf_file.name,
                        pdf_file.read(),
                        "application/pdf",
                    )
                }
                resp = requests.post(PDF_BACKEND_URL, files=files, timeout=900)
                resp.raise_for_status()
                data = resp.json()
                
                if isinstance(data, dict) and "text" in data:
                    st.session_state["pdf_text"] = data["text"]
                elif isinstance(data, dict) and "pdf_text" in data:
                    st.session_state["pdf_text"] = data["pdf_text"]
                elif isinstance(data, str):
                    st.session_state["pdf_text"] = data
                else:
                    st.session_state["pdf_text"] = None
                    st.error("Unexpected response from pdf-to-text backend.")
                st.session_state["pdf_uploaded_name"] = pdf_file.name

                if st.session_state["pdf_text"]:
                    st.success("DRHP PDF uploaded successfully.")
                else:
                    st.warning("Backend returned empty text for the PDF.")
            except Exception as e:
                st.error(f"Error calling pdf-to-text backend: {e}")
    else:
        st.session_state["pdf_text"] = None
        st.session_state["pdf_uploaded_name"] = None

    # st.markdown("---")

    # MAIN ACTION BUTTON
    if st.button("Summarize and Populate Template"):
        if not st.session_state["template_json"]:
            st.error("Template JSON not available. Upload DOCX first.")
        elif not st.session_state["pdf_text"]:
            st.error("PDF text not available. Upload PDF first.")
        else:
            # Placeholders for progress & Status
            progress_bar = st.progress(0)
            status_text = st.empty()
            try:
                status_text.info("Step 1/3: Preparing Data for AI Model...")
                time.sleep(2)
                # ---- Step 1: Prepare Inputs ----
                txt_bytes = st.session_state["pdf_text"].encode("utf-8")
                schema_str = json.dumps(
                    st.session_state["template_json"],
                    separators=(",", ":"),
                    ensure_ascii=False,
                )
                print(schema_str)
                progress_bar.progress(20)
                files = {
                        "txt_file": ("input.txt", txt_bytes, "text/plain"),
                    }
                data = {
                        "input_schema": schema_str,
                        "chunk_size":10000,
                        "web_sections":"5,6,7,8,11,14"
                    }
                # ---- Step 2: Call LLM (/process) ----
                status_text.info("Step 2/3: AI is analyzing the document...")
                with st.spinner("AI is extracting information from the PDF..."):
                    resp = requests.post(LLM_API_URL, files=files, data=data, timeout=6000)
                resp.raise_for_status()
                resp_json = resp.json()
                print("\n\n")
                print(resp_json)

                if isinstance(resp_json, dict) and "filled_json" in resp_json:
                    filled_json = resp_json["filled_json"]
                else:
                    filled_json = resp_json
                progress_bar.progress(65)

                # ---- Step 3: Call /fill-html with filled_json + built-in HTML template ----
                status_text.info("Step 3/3: Generating formatted IM in HTML...")
                time.sleep(1)
                with st.spinner("Generating formatted Output..."):
                    print("=============")
                    print(type(filled_json))
                    print("===========================================")
                # Adjust payload structure as per your actual /fill-html contract
                    # payload = {
                    #     "data": filled_json,
                    #     "template_html": HTML_TEMPLATE,
                    # }
                    
                    payload = {
                        "template_html": HTML_TEMPLATE,
                        "data": filled_json["extracted_json"]
                    }
                    html_resp = requests.post(FILL_HTML_URL, json=payload, timeout=600)
                    # html_resp = requests.post(FILL_HTML_URL, data = {"data":json.dumps(payload["data"]),"template_html":payload["template_html"]}, timeout=600)
                    html_resp.raise_for_status()

                    try:
                        html_data = html_resp.json()
                        if isinstance(html_data, dict) and "html" in html_data:
                            html = html_data["html"]
                        else:
                            html = str(html_data)
                    except ValueError:
                        html = html_resp.text

                    st.session_state["filled_html"] = html
                    progress_bar.progress(100)
                    status_text.success("Done! Document processed successfully!")
                    # st.success("HTML generated.")
            except Exception as e:
                st.error(f"Error during extraction/HTML generation: {str(e)}")
            finally:
                progress_bar.empty()    

# -------- RIGHT: Render full HTML and show download buttons --------
with right:
    if st.session_state.get("filled_html"):
        # 1) Render the complete HTML directly
        st.components.v1.html(st.session_state["filled_html"], height=700, scrolling=True) # type: ignore
        st.markdown(
            """
            <style>
            iframe{
            border: 1px solid #85929E !important;
            border-radius: 6px !important;
            box-shadow: 0 4px 12px rgba(0,0,0,0.15);
        
            }
            div[data-testId="stVerticalBlock"]{
            margin-top: 15px !important;
            }
            </style>
            """,
            unsafe_allow_html=True
        )
        c1,c2 = st.columns([2,2])
        with c1:
            import requests




            try:
                payload = {
                    "html": st.session_state["filled_html"],
                    "file_name": "im_output.docx"
                }

                resp = requests.post(
                    HTML_TO_DOCX_URL,
                    json=payload,
                    timeout=120
                )
                # Form-data
                # resp = requests.post(
                #     HTML_TO_DOCX_URL,
                #     data={"html": st.session_state["filled_html"]},
                #     timeout=120
                # )

                resp.raise_for_status()

                docx_bytes = resp.content

                st.download_button(
                    "⬇️ Download as DOCX",
                    data=docx_bytes,
                    file_name="im_output.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                )

            except Exception as e:
                st.error(f"Error creating DOCX: {e}")
                
    
    else:
        # st.info("Upload PDF, then click 'Extract Information with AI' to see the output here.")
        pass

st.markdown("""
<style>
/* 1️⃣ Hide original button text by making font-size 0 */ 
div[data-testid="stFileUploader"] button[kind="secondary"]{
    font-size: 0 !important;
}
 
/* 2️⃣ Add custom text via ::after */ 
div[data-testid="stFileUploader"] button[kind="secondary"]::after {
    content: "Upload File";
    font-size: 16px !important;
    display: inline-block;
}
</style>
""", unsafe_allow_html=True)
