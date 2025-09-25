"""
PWD Tools - Reorganized Application
Professional PWD-themed tools with enhanced styling
"""

import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, date
import os
import sys

# Add tools directories to path
sys.path.append(os.path.join(os.path.dirname(__file__), 'tools'))
sys.path.append(os.path.join(os.path.dirname(__file__), 'utils'))

# Import tool modules
from tools.financial import bill_note_sheet, emd_refund, security_refund, financial_progress, hindi_bill_generator
from tools.calculation import delay_calculator, stamp_duty_calculator, deductions_table, excel_emd_processor
from tools.reports import bill_deviation_generator, financial_analysis

def set_page_config():
    """Configure page settings with PWD theme"""
    st.set_page_config(
        page_title="PWD Tools - Infrastructure Management Suite",
        page_icon="🏗️",
        layout="wide",
        initial_sidebar_state="expanded"
    )

def render_header():
    """Render main application header"""
    st.markdown("""
        <div style="text-align: center; background: linear-gradient(135deg, #FF6B35 0%, #004E89 100%); 
                    padding: 2rem; border-radius: 10px; color: white; margin-bottom: 2rem; 
                    box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);">
            <h1 style="margin: 0; font-size: 2.5rem; font-weight: 700;">🏗️ PWD Tools - Infrastructure Management Suite</h1>
            <p style="margin: 0.5rem 0 0 0; font-size: 1.1rem; opacity: 0.9;">
                Comprehensive tools for Public Works Department operations with enhanced professional styling
            </p>
        </div>
    """, unsafe_allow_html=True)

def render_sidebar():
    """Render sidebar with navigation and statistics"""
    st.sidebar.title("🏗️ PWD Navigation")
    
    # Tool categories
    st.sidebar.header("📊 Tool Categories")
    
    # Statistics
    st.sidebar.markdown("### 📈 Quick Stats")
    
    col1, col2 = st.sidebar.columns(2)
    with col1:
        st.sidebar.metric("Total Tools", "11")
    
    with col2:
        st.sidebar.metric("Categories", "3")
    
    # Navigation menu
    st.sidebar.header("🔧 Quick Access")
    
    if st.sidebar.button("📋 Financial Tools", use_container_width=True):
        st.session_state.selected_category = "financial"
        st.rerun()
    
    if st.sidebar.button("🧮 Calculation Tools", use_container_width=True):
        st.session_state.selected_category = "calculation"
        st.rerun()
    
    if st.sidebar.button("📊 Report Tools", use_container_width=True):
        st.session_state.selected_category = "reports"
        st.rerun()
    
    if st.sidebar.button("🏠 Dashboard", use_container_width=True):
        st.session_state.selected_category = "dashboard"
        st.rerun()
    
    # User info
    st.sidebar.markdown("---")
    st.sidebar.markdown("### 👤 User Information")
    st.sidebar.info("**Prepared for:**\nMrs. Premlata Jain, AAO\nPWD Udaipur")
    
    # Version info
    st.sidebar.markdown("---")
    st.sidebar.markdown("**Version:** 2.0.0")
    st.sidebar.markdown("**Last Updated:** September 2024")

def render_dashboard():
    """Render main dashboard with tool categories"""
    
    # Introduction
    st.markdown("### 🎯 Welcome to PWD Tools Suite")
    st.markdown("""
    This comprehensive suite provides essential tools for Public Works Department operations, 
    organized into three main categories for better workflow management.
    """)
    
    # Tool categories
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.markdown("""
            <div style="background: white; border: 2px solid #e9ecef; border-radius: 10px; 
                        padding: 1.5rem; margin: 1rem 0; box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1); 
                        transition: all 0.3s ease;">
                <div style="display: flex; align-items: center; margin-bottom: 1rem;">
                    <span style="font-size: 2rem; margin-right: 1rem;">💰</span>
                    <h3 style="font-size: 1.5rem; font-weight: 600; color: #004E89; margin: 0;">Financial Tools</h3>
                </div>
                <p style="color: #6C757D; margin-bottom: 1rem; font-style: italic;">
                    Essential financial management and documentation tools
                </p>
                <div>
                    <strong>5 Tools Available:</strong><br>
                    • Bill Note Sheet<br>
                    • EMD Refund<br>
                    • Security Refund<br>
                    • Financial Progress<br>
                    • Hindi Bill Generator
                </div>
            </div>
        """, unsafe_allow_html=True)
        
        if st.button("Access Financial Tools", key="btn_financial", use_container_width=True):
            st.session_state.selected_category = "financial"
            st.rerun()
    
    with col2:
        st.markdown("""
            <div style="background: white; border: 2px solid #e9ecef; border-radius: 10px; 
                        padding: 1.5rem; margin: 1rem 0; box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1); 
                        transition: all 0.3s ease;">
                <div style="display: flex; align-items: center; margin-bottom: 1rem;">
                    <span style="font-size: 2rem; margin-right: 1rem;">🧮</span>
                    <h3 style="font-size: 1.5rem; font-weight: 600; color: #004E89; margin: 0;">Calculation Tools</h3>
                </div>
                <p style="color: #6C757D; margin-bottom: 1rem; font-style: italic;">
                    Advanced calculation and processing utilities
                </p>
                <div>
                    <strong>4 Tools Available:</strong><br>
                    • Delay Calculator<br>
                    • Stamp Duty Calculator<br>
                    • Deductions Table<br>
                    • Excel EMD Processor
                </div>
            </div>
        """, unsafe_allow_html=True)
        
        if st.button("Access Calculation Tools", key="btn_calculation", use_container_width=True):
            st.session_state.selected_category = "calculation"
            st.rerun()
    
    with col3:
        st.markdown("""
            <div style="background: white; border: 2px solid #e9ecef; border-radius: 10px; 
                        padding: 1.5rem; margin: 1rem 0; box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1); 
                        transition: all 0.3s ease;">
                <div style="display: flex; align-items: center; margin-bottom: 1rem;">
                    <span style="font-size: 2rem; margin-right: 1rem;">📊</span>
                    <h3 style="font-size: 1.5rem; font-weight: 600; color: #004E89; margin: 0;">Report Tools</h3>
                </div>
                <p style="color: #6C757D; margin-bottom: 1rem; font-style: italic;">
                    Comprehensive reporting and analysis solutions
                </p>
                <div>
                    <strong>2 Tools Available:</strong><br>
                    • Bill & Deviation Generator<br>
                    • Financial Analysis
                </div>
            </div>
        """, unsafe_allow_html=True)
        
        if st.button("Access Report Tools", key="btn_reports", use_container_width=True):
            st.session_state.selected_category = "reports"
            st.rerun()

def render_financial_tools():
    """Render financial tools section"""
    st.markdown("## 💰 Financial Tools")
    st.markdown("Essential financial management and documentation tools for PWD operations")
    
    # Tool selection
    tool_option = st.selectbox(
        "Select a Financial Tool:",
        ["Select Tool", "Bill Note Sheet", "EMD Refund", "Security Refund", "Financial Progress", "Hindi Bill Generator"]
    )
    
    if tool_option == "Bill Note Sheet":
        bill_note_sheet.main()
    elif tool_option == "EMD Refund":
        emd_refund.main()
    elif tool_option == "Security Refund":
        security_refund.main()
    elif tool_option == "Financial Progress":
        financial_progress.main()
    elif tool_option == "Hindi Bill Generator":
        hindi_bill_generator.main()
    elif tool_option == "Select Tool":
        st.info("Please select a financial tool from the dropdown above to begin.")

def render_calculation_tools():
    """Render calculation tools section"""
    st.markdown("## 🧮 Calculation Tools")
    st.markdown("Advanced calculation and processing utilities for engineering and administrative tasks")
    
    # Tool selection
    tool_option = st.selectbox(
        "Select a Calculation Tool:",
        ["Select Tool", "Delay Calculator", "Stamp Duty Calculator", "Deductions Table", "Excel EMD Processor"]
    )
    
    if tool_option == "Delay Calculator":
        delay_calculator.main()
    elif tool_option == "Stamp Duty Calculator":
        stamp_duty_calculator.main()
    elif tool_option == "Deductions Table":
        deductions_table.main()
    elif tool_option == "Excel EMD Processor":
        excel_emd_processor.main()
    elif tool_option == "Select Tool":
        st.info("Please select a calculation tool from the dropdown above to begin.")

def render_report_tools():
    """Render report tools section"""
    st.markdown("## 📊 Report Tools")
    st.markdown("Comprehensive reporting and analysis solutions for project management")
    
    # Tool selection
    tool_option = st.selectbox(
        "Select a Report Tool:",
        ["Select Tool", "Bill & Deviation Generator", "Financial Analysis"]
    )
    
    if tool_option == "Bill & Deviation Generator":
        bill_deviation_generator.main()
    elif tool_option == "Financial Analysis":
        financial_analysis.main()
    elif tool_option == "Select Tool":
        st.info("Please select a report tool from the dropdown above to begin.")

def render_footer():
    """Render application footer"""
    st.markdown("""
        <div style="background-color: #004E89; color: white; text-align: center; 
                    padding: 2rem; margin-top: 3rem; border-radius: 10px;">
            <h4 style="margin: 0 0 0.5rem 0; color: #FF6B35;">🏗️ PWD Tools - Infrastructure Management Suite</h4>
            <p style="margin: 0; opacity: 0.9;">Prepared for Mrs. Premlata Jain, AAO, PWD Udaipur</p>
            <p style="margin: 0; opacity: 0.9;">Enhanced with professional PWD-themed styling | Version 2.0.0</p>
        </div>
    """, unsafe_allow_html=True)

def main():
    """Main application function"""
    # Set page configuration
    set_page_config()
    
    # Initialize session state
    if 'selected_category' not in st.session_state:
        st.session_state.selected_category = "dashboard"
    
    # Render header
    render_header()
    
    # Render sidebar
    render_sidebar()
    
    # Render main content based on selected category
    if st.session_state.selected_category == "financial":
        render_financial_tools()
    elif st.session_state.selected_category == "calculation":
        render_calculation_tools()
    elif st.session_state.selected_category == "reports":
        render_report_tools()
    else:
        render_dashboard()
    
    # Render footer
    render_footer()

if __name__ == "__main__":
    main()
