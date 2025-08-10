#!/usr/bin/env python3
"""
GUI Table Scraper - A user-friendly interface for the enhanced table scraper
Features:
- Easy URL input
- Output filename selection
- File format selection
- Screenshot toggle
- Real-time progress updates
- Log viewer
"""

import json
import logging
import os
import threading
import time
import tkinter as tk
from pathlib import Path
from tkinter import ttk, filedialog, messagebox, scrolledtext
from typing import List, Optional, Tuple

import pandas as pd
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.common.exceptions import TimeoutException, NoSuchElementException, WebDriverException
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait


class TableScraperConfig:
    """Configuration class for the table scraper"""
    def __init__(self):
        self.url = ""
        self.output_file = "scraped_tables"
        self.output_format = "excel"
        self.headless = True
        self.timeout = 10
        self.max_retries = 3
        self.retry_delay = 2
        self.take_screenshots = False
        self.chrome_binary_path = None
        self.max_pages = None
        self.output_directory = "output"
        self.next_button_selectors = [
            "//a[contains(text(), 'Next')]",
            "//button[contains(text(), 'Next')]",
            "//a[contains(@class, 'next')]",
            "//button[contains(@class, 'next')]"
        ]


class GUILogHandler(logging.Handler):
    """Custom logging handler to display logs in GUI"""
    
    def __init__(self, text_widget):
        super().__init__()
        self.text_widget = text_widget
        
    def emit(self, record):
        msg = self.format(record)
        # Use after_idle to ensure thread safety
        self.text_widget.after_idle(self._append_log, msg)
        
    def _append_log(self, msg):
        self.text_widget.insert(tk.END, msg + '\n')
        self.text_widget.see(tk.END)


class EnhancedTableScraper:
    """Enhanced table scraper with GUI integration"""
    
    def __init__(self, config: TableScraperConfig, progress_callback=None, status_callback=None):
        self.config = config
        self.driver = None
        self.wait = None
        self.all_data = []
        self.progress_callback = progress_callback
        self.status_callback = status_callback
        self.is_running = False
        self.should_stop = False
        
    def setup_logging(self, log_widget=None):
        """Setup logging configuration with optional GUI handler"""
        # Clear existing handlers
        for handler in logging.root.handlers[:]:
            logging.root.removeHandler(handler)
            
        # Setup formatters
        formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
        
        # File handler
        file_handler = logging.FileHandler('scraper.log')
        file_handler.setFormatter(formatter)
        
        # Console handler
        console_handler = logging.StreamHandler()
        console_handler.setFormatter(formatter)
        
        # Setup root logger
        logging.basicConfig(level=logging.INFO, handlers=[file_handler, console_handler])
        
        # Add GUI handler if provided
        if log_widget:
            gui_handler = GUILogHandler(log_widget)
            gui_handler.setFormatter(logging.Formatter('%(levelname)s: %(message)s'))
            logging.getLogger().addHandler(gui_handler)
            
        self.logger = logging.getLogger(__name__)
        
    def update_status(self, message):
        """Update status in GUI"""
        if self.status_callback:
            self.status_callback(message)
        self.logger.info(message)
        
    def update_progress(self, current, total):
        """Update progress in GUI"""
        if self.progress_callback:
            self.progress_callback(current, total)
            
    def setup_driver(self) -> bool:
        """Setup Chrome WebDriver with options"""
        try:
            self.update_status("Initializing browser...")
            options = Options()
            if self.config.headless:
                options.add_argument("--headless")
            
            options.add_argument("--no-sandbox")
            options.add_argument("--disable-dev-shm-usage")
            options.add_argument("--disable-gpu")
            options.add_argument("--window-size=1920,1080")
            
            if self.config.chrome_binary_path:
                options.binary_location = self.config.chrome_binary_path
                
            self.driver = webdriver.Chrome(options=options)
            self.wait = WebDriverWait(self.driver, self.config.timeout)
            self.update_status("Browser initialized successfully")
            return True
            
        except Exception as e:
            self.update_status(f"Failed to initialize browser: {e}")
            return False
    
    def load_page_with_retry(self, url: str) -> bool:
        """Load page with retry mechanism"""
        for attempt in range(self.config.max_retries):
            if self.should_stop:
                return False
                
            try:
                self.update_status(f"Loading page (attempt {attempt + 1}): {url}")
                self.driver.get(url)
                
                self.wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))
                self.wait.until(EC.presence_of_element_located((By.TAG_NAME, "table")))
                
                self.update_status("Page loaded successfully")
                return True
                
            except TimeoutException:
                self.update_status(f"Timeout loading page (attempt {attempt + 1})")
                if attempt < self.config.max_retries - 1:
                    time.sleep(self.config.retry_delay)
                    
            except WebDriverException as e:
                self.update_status(f"WebDriver error (attempt {attempt + 1}): {e}")
                if attempt < self.config.max_retries - 1:
                    time.sleep(self.config.retry_delay)
                    
        self.update_status("Failed to load page after all retry attempts")
        return False
    
    def extract_tables_from_page(self, page_num: int) -> List[pd.DataFrame]:
        """Extract all tables from current page"""
        try:
            soup = BeautifulSoup(self.driver.page_source, "html.parser")
            tables = soup.find_all("table")
            
            if not tables:
                self.update_status(f"No tables found on page {page_num}")
                return []
            
            page_dataframes = []
            for table_idx, table in enumerate(tables):
                try:
                    df_list = pd.read_html(str(table))
                    if df_list:
                        df = df_list[0]
                        
                        df['source_page'] = page_num
                        df['source_url'] = self.driver.current_url
                        df['scraped_at'] = pd.Timestamp.now()
                        
                        page_dataframes.append(df)
                        self.update_status(f"Extracted table {table_idx + 1} from page {page_num} ({len(df)} rows)")
                        
                except Exception as e:
                    self.update_status(f"Error parsing table {table_idx + 1} on page {page_num}: {e}")
                    
            return page_dataframes
            
        except Exception as e:
            self.update_status(f"Error extracting tables from page {page_num}: {e}")
            return []
    
    def take_screenshot(self, page_num: int):
        """Take screenshot of current page"""
        if self.config.take_screenshots and not self.should_stop:
            try:
                screenshot_dir = Path(self.config.output_directory)
                screenshot_dir.mkdir(exist_ok=True)
                screenshot_path = screenshot_dir / f"screenshot_page_{page_num}.png"
                self.driver.save_screenshot(str(screenshot_path))
                self.update_status(f"Screenshot saved: {screenshot_path}")
            except Exception as e:
                self.update_status(f"Failed to take screenshot for page {page_num}: {e}")
    
    def find_next_button(self) -> Optional[object]:
        """Find next button using multiple selectors"""
        for selector in self.config.next_button_selectors:
            try:
                element = self.driver.find_element(By.XPATH, selector)
                
                if (element.is_enabled() and 
                    element.is_displayed() and 
                    "disabled" not in element.get_attribute("class").lower()):
                    return element
                    
            except NoSuchElementException:
                continue
                
        return None
    
    def has_next_page(self) -> Tuple[bool, Optional[object]]:
        """Check if there's a next page available"""
        next_button = self.find_next_button()
        
        if next_button is None:
            return False, None
            
        try:
            button_text = next_button.text.lower()
            if any(word in button_text for word in ['disabled', 'last', 'end']):
                return False, None
                
            if next_button.get_attribute("disabled"):
                return False, None
                
            return True, next_button
            
        except Exception as e:
            self.update_status(f"Error checking next page availability: {e}")
            return False, None
    
    def scrape_all_pages(self) -> bool:
        """Main scraping logic for all pages"""
        self.is_running = True
        self.should_stop = False
        
        if not self.load_page_with_retry(self.config.url):
            self.is_running = False
            return False
            
        page_num = 1
        
        while not self.should_stop:
            self.update_status(f"Processing page {page_num}")
            self.update_progress(page_num, page_num + 1)  # Indeterminate progress
            
            self.take_screenshot(page_num)
            
            page_tables = self.extract_tables_from_page(page_num)
            self.all_data.extend(page_tables)
            
            if self.config.max_pages and page_num >= self.config.max_pages:
                self.update_status(f"Reached maximum pages limit: {self.config.max_pages}")
                break
            
            has_next, next_button = self.has_next_page()
            if not has_next:
                self.update_status("No more pages found")
                break
                
            try:
                next_button.click()
                
                self.wait.until(lambda driver: str(page_num + 1) in driver.current_url or 
                               driver.find_element(By.TAG_NAME, "table"))
                
                page_num += 1
                time.sleep(1)
                
            except Exception as e:
                self.update_status(f"Error navigating to next page: {e}")
                break
                
        self.update_status(f"Scraping completed. Processed {page_num} pages, extracted {len(self.all_data)} tables")
        self.is_running = False
        return True
    
    def save_data(self):
        """Save extracted data in specified format(s)"""
        if not self.all_data:
            self.update_status("No data to save")
            return False
            
        final_df = pd.concat(self.all_data, ignore_index=True)
        self.update_status(f"Combined data: {len(final_df)} total rows")
        
        Path(self.config.output_directory).mkdir(exist_ok=True)
        
        try:
            if self.config.output_format in ["excel", "all"]:
                self.save_excel(final_df)
                
            if self.config.output_format in ["csv", "all"]:
                self.save_csv(final_df)
                
            if self.config.output_format in ["json", "all"]:
                self.save_json(final_df)
                
            return True
        except Exception as e:
            self.update_status(f"Error saving data: {e}")
            return False
    
    def save_excel(self, df: pd.DataFrame):
        """Save data to Excel with multiple sheets"""
        excel_path = Path(self.config.output_directory) / f"{self.config.output_file}.xlsx"
        
        with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name="All_Data", index=False)
            
            for page_num in sorted(df['source_page'].unique()):
                page_data = df[df['source_page'] == page_num]
                sheet_name = f"Page_{page_num}"
                page_data.to_excel(writer, sheet_name=sheet_name, index=False)
                
        self.update_status(f"Excel file saved: {excel_path}")
    
    def save_csv(self, df: pd.DataFrame):
        """Save data to CSV"""
        csv_path = Path(self.config.output_directory) / f"{self.config.output_file}.csv"
        df.to_csv(csv_path, index=False)
        self.update_status(f"CSV file saved: {csv_path}")
    
    def save_json(self, df: pd.DataFrame):
        """Save data to JSON"""
        json_path = Path(self.config.output_directory) / f"{self.config.output_file}.json"
        
        df_json = df.copy()
        df_json['scraped_at'] = df_json['scraped_at'].astype(str)
        
        with open(json_path, 'w', encoding='utf-8') as f:
            json.dump(df_json.to_dict('records'), f, indent=2, ensure_ascii=False)
            
        self.update_status(f"JSON file saved: {json_path}")
    
    def stop_scraping(self):
        """Stop the scraping process"""
        self.should_stop = True
        self.update_status("Stopping scraper...")
    
    def cleanup(self):
        """Cleanup resources"""
        if self.driver:
            self.driver.quit()
            self.update_status("Browser closed")


class TableScraperGUI:
    """Main GUI application for the table scraper"""
    
    def __init__(self, root):
        self.root = root
        self.root.title("Table Scraper - GUI Version")
        self.root.geometry("800x700")
        
        self.scraper = None
        self.scraper_thread = None
        
        self.setup_ui()
        
    def setup_ui(self):
        """Setup the user interface"""
        # Main notebook for tabs
        notebook = ttk.Notebook(self.root)
        notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Configuration tab
        config_frame = ttk.Frame(notebook)
        notebook.add(config_frame, text="Configuration")
        self.setup_config_tab(config_frame)
        
        # Progress tab
        progress_frame = ttk.Frame(notebook)
        notebook.add(progress_frame, text="Progress & Logs")
        self.setup_progress_tab(progress_frame)
        
        # About tab
        about_frame = ttk.Frame(notebook)
        notebook.add(about_frame, text="About")
        self.setup_about_tab(about_frame)
    
    def setup_config_tab(self, parent):
        """Setup configuration tab"""
        # URL Configuration
        url_frame = ttk.LabelFrame(parent, text="URL Configuration", padding=10)
        url_frame.pack(fill=tk.X, padx=10, pady=5)
        
        ttk.Label(url_frame, text="Target URL:").pack(anchor=tk.W)
        self.url_var = tk.StringVar(value="https://www.sensitron.com/index.php?action=sic_fet")
        self.url_entry = ttk.Entry(url_frame, textvariable=self.url_var, width=80)
        self.url_entry.pack(fill=tk.X, pady=(0, 10))
        
        # Output Configuration
        output_frame = ttk.LabelFrame(parent, text="Output Configuration", padding=10)
        output_frame.pack(fill=tk.X, padx=10, pady=5)
        
        # Output filename
        ttk.Label(output_frame, text="Output Filename (without extension):").pack(anchor=tk.W)
        self.filename_var = tk.StringVar(value="scraped_tables")
        ttk.Entry(output_frame, textvariable=self.filename_var, width=40).pack(anchor=tk.W, pady=(0, 10))
        
        # Output directory
        dir_frame = ttk.Frame(output_frame)
        dir_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(dir_frame, text="Output Directory:").pack(anchor=tk.W)
        dir_inner_frame = ttk.Frame(dir_frame)
        dir_inner_frame.pack(fill=tk.X)
        
        self.output_dir_var = tk.StringVar(value="output")
        ttk.Entry(dir_inner_frame, textvariable=self.output_dir_var, width=50).pack(side=tk.LEFT, fill=tk.X, expand=True)
        ttk.Button(dir_inner_frame, text="Browse", command=self.browse_output_dir).pack(side=tk.RIGHT, padx=(5, 0))
        
        # File format
        ttk.Label(output_frame, text="Output Format:").pack(anchor=tk.W)
        self.format_var = tk.StringVar(value="excel")
        format_frame = ttk.Frame(output_frame)
        format_frame.pack(anchor=tk.W, pady=(0, 10))
        
        formats = [("Excel (.xlsx)", "excel"), ("CSV (.csv)", "csv"), ("JSON (.json)", "json"), ("All formats", "all")]
        for text, value in formats:
            ttk.Radiobutton(format_frame, text=text, variable=self.format_var, value=value).pack(anchor=tk.W)
        
        # Options Configuration
        options_frame = ttk.LabelFrame(parent, text="Scraping Options", padding=10)
        options_frame.pack(fill=tk.X, padx=10, pady=5)
        
        # Screenshots option
        self.screenshots_var = tk.BooleanVar()
        ttk.Checkbutton(options_frame, text="Take screenshots of each page", 
                       variable=self.screenshots_var).pack(anchor=tk.W, pady=2)
        
        # Headless option
        self.headless_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(options_frame, text="Run browser in headless mode (recommended)", 
                       variable=self.headless_var).pack(anchor=tk.W, pady=2)
        
        # Max pages
        max_pages_frame = ttk.Frame(options_frame)
        max_pages_frame.pack(anchor=tk.W, pady=2)
        
        self.max_pages_enabled = tk.BooleanVar()
        ttk.Checkbutton(max_pages_frame, text="Limit maximum pages:", 
                       variable=self.max_pages_enabled).pack(side=tk.LEFT)
        
        self.max_pages_var = tk.StringVar(value="10")
        max_pages_spinbox = ttk.Spinbox(max_pages_frame, from_=1, to=1000, width=10, 
                                       textvariable=self.max_pages_var)
        max_pages_spinbox.pack(side=tk.LEFT, padx=(5, 0))
        
        # Advanced Configuration
        advanced_frame = ttk.LabelFrame(parent, text="Advanced Settings", padding=10)
        advanced_frame.pack(fill=tk.X, padx=10, pady=5)
        
        # Timeout
        timeout_frame = ttk.Frame(advanced_frame)
        timeout_frame.pack(anchor=tk.W, pady=2)
        
        ttk.Label(timeout_frame, text="Timeout (seconds):").pack(side=tk.LEFT)
        self.timeout_var = tk.StringVar(value="10")
        ttk.Spinbox(timeout_frame, from_=5, to=60, width=10, 
                   textvariable=self.timeout_var).pack(side=tk.LEFT, padx=(5, 0))
        
        # Chrome binary path
        chrome_frame = ttk.Frame(advanced_frame)
        chrome_frame.pack(fill=tk.X, pady=2)
        
        ttk.Label(chrome_frame, text="Chrome Binary Path (optional):").pack(anchor=tk.W)
        chrome_inner_frame = ttk.Frame(chrome_frame)
        chrome_inner_frame.pack(fill=tk.X)
        
        self.chrome_path_var = tk.StringVar()
        ttk.Entry(chrome_inner_frame, textvariable=self.chrome_path_var).pack(side=tk.LEFT, fill=tk.X, expand=True)
        ttk.Button(chrome_inner_frame, text="Browse", command=self.browse_chrome_binary).pack(side=tk.RIGHT, padx=(5, 0))
        
        # Control buttons
        button_frame = ttk.Frame(parent)
        button_frame.pack(fill=tk.X, padx=10, pady=10)
        
        self.start_button = ttk.Button(button_frame, text="🚀 Start Scraping", command=self.start_scraping)
        self.start_button.pack(side=tk.LEFT, padx=(0, 5))
        
        self.stop_button = ttk.Button(button_frame, text="⏹️ Stop Scraping", command=self.stop_scraping, state=tk.DISABLED)
        self.stop_button.pack(side=tk.LEFT, padx=5)
        
        ttk.Button(button_frame, text="📁 Open Output Folder", command=self.open_output_folder).pack(side=tk.RIGHT)
    
    def setup_progress_tab(self, parent):
        """Setup progress and logs tab"""
        # Status
        status_frame = ttk.LabelFrame(parent, text="Status", padding=10)
        status_frame.pack(fill=tk.X, padx=10, pady=5)
        
        self.status_var = tk.StringVar(value="Ready to start scraping...")
        ttk.Label(status_frame, textvariable=self.status_var, font=('TkDefaultFont', 10, 'bold')).pack()
        
        # Progress bar
        progress_frame = ttk.LabelFrame(parent, text="Progress", padding=10)
        progress_frame.pack(fill=tk.X, padx=10, pady=5)
        
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(progress_frame, variable=self.progress_var, mode='indeterminate')
        self.progress_bar.pack(fill=tk.X)
        
        # Logs
        logs_frame = ttk.LabelFrame(parent, text="Logs", padding=10)
        logs_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        self.log_text = scrolledtext.ScrolledText(logs_frame, height=20, state=tk.DISABLED)
        self.log_text.pack(fill=tk.BOTH, expand=True)
        
        # Clear logs button
        ttk.Button(logs_frame, text="Clear Logs", command=self.clear_logs).pack(pady=(5, 0))
    
    def setup_about_tab(self, parent):
        """Setup about tab"""
        about_text = """
        🕷️ Table Scraper GUI v1.0
        
        A powerful web scraper for extracting tables from paginated websites.
        
        Features:
        • 🌐 Easy URL input and configuration
        • 📊 Multiple output formats (Excel, CSV, JSON)
        • 📸 Optional screenshot capture
        • 🔄 Robust pagination handling
        • 📝 Real-time progress monitoring
        • 🛡️ Error handling and retry mechanisms
        
        How to use:
        1. Enter the target URL in the Configuration tab
        2. Choose your output preferences
        3. Configure scraping options as needed
        4. Click "Start Scraping" and monitor progress
        
        Requirements:
        • Python packages: selenium, beautifulsoup4, pandas, openpyxl
        • Chrome browser and ChromeDriver
        
        Tips:
        • Use headless mode for better performance
        • Enable screenshots for debugging
        • Set reasonable limits for large sites
        • Check logs for detailed information
        
        Happy scraping! 🎉
        """
        
        text_widget = scrolledtext.ScrolledText(parent, wrap=tk.WORD, state=tk.DISABLED)
        text_widget.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        text_widget.config(state=tk.NORMAL)
        text_widget.insert(tk.END, about_text)
        text_widget.config(state=tk.DISABLED)
    
    def browse_output_dir(self):
        """Browse for output directory"""
        directory = filedialog.askdirectory(initialdir=self.output_dir_var.get())
        if directory:
            self.output_dir_var.set(directory)
    
    def browse_chrome_binary(self):
        """Browse for Chrome binary"""
        filename = filedialog.askopenfilename(
            title="Select Chrome Binary",
            filetypes=[("Executable files", "*.exe"), ("All files", "*.*")]
        )
        if filename:
            self.chrome_path_var.set(filename)
    
    def create_config_from_gui(self) -> TableScraperConfig:
        """Create configuration object from GUI inputs"""
        config = TableScraperConfig()
        
        config.url = self.url_var.get().strip()
        config.output_file = self.filename_var.get().strip() or "scraped_tables"
        config.output_directory = self.output_dir_var.get().strip() or "output"
        config.output_format = self.format_var.get()
        config.take_screenshots = self.screenshots_var.get()
        config.headless = self.headless_var.get()
        config.timeout = int(self.timeout_var.get())
        
        if self.max_pages_enabled.get():
            config.max_pages = int(self.max_pages_var.get())
        
        chrome_path = self.chrome_path_var.get().strip()
        if chrome_path:
            config.chrome_binary_path = chrome_path
            
        return config
    
    def validate_inputs(self) -> bool:
        """Validate user inputs"""
        url = self.url_var.get().strip()
        if not url:
            messagebox.showerror("Error", "Please enter a valid URL")
            return False
            
        if not url.startswith(('http://', 'https://')):
            messagebox.showerror("Error", "URL must start with http:// or https://")
            return False
            
        filename = self.filename_var.get().strip()
        if not filename:
            messagebox.showerror("Error", "Please enter an output filename")
            return False
            
        try:
            int(self.timeout_var.get())
        except ValueError:
            messagebox.showerror("Error", "Timeout must be a valid number")
            return False
            
        if self.max_pages_enabled.get():
            try:
                max_pages = int(self.max_pages_var.get())
                if max_pages < 1:
                    raise ValueError()
            except ValueError:
                messagebox.showerror("Error", "Maximum pages must be a positive number")
                return False
                
        return True
    
    def update_status(self, message):
        """Update status label"""
        self.status_var.set(message)
        self.root.update_idletasks()
    
    def update_progress(self, current, total):
        """Update progress bar"""
        # For indeterminate progress, just pulse the bar
        if not self.progress_bar.cget('mode') == 'indeterminate':
            self.progress_bar.config(mode='indeterminate')
        self.progress_bar.start(10)
    
    def clear_logs(self):
        """Clear the log text widget"""
        self.log_text.config(state=tk.NORMAL)
        self.log_text.delete(1.0, tk.END)
        self.log_text.config(state=tk.DISABLED)
    
    def start_scraping(self):
        """Start the scraping process"""
        if not self.validate_inputs():
            return
            
        # Update UI state
        self.start_button.config(state=tk.DISABLED)
        self.stop_button.config(state=tk.NORMAL)
        self.progress_bar.start(10)
        
        # Clear previous logs
        self.clear_logs()
        
        # Create configuration
        config = self.create_config_from_gui()
        
        # Create scraper with callbacks
        self.scraper = EnhancedTableScraper(
            config,
            progress_callback=self.update_progress,
            status_callback=self.update_status
        )
        
        # Setup logging with GUI handler
        self.scraper.setup_logging(self.log_text)
        
        # Start scraping in separate thread
        self.scraper_thread = threading.Thread(target=self.run_scraper, daemon=True)
        self.scraper_thread.start()
    
    def run_scraper(self):
        """Run scraper in separate thread"""
        try:
            if not self.scraper.setup_driver():
                self.scraping_finished(False)
                return
                
            success = self.scraper.scrape_all_pages()
            
            if success and self.scraper.all_data:
                save_success = self.scraper.save_data()
                self.scraping_finished(save_success)
            else:
                self.scraping_finished(False)
                
        except Exception as e:
            self.update_status(f"Unexpected error: {e}")
            self.scraping_finished(False)
        finally:
            self.scraper.cleanup()
    
    def scraping_finished(self, success):
        """Handle scraping completion"""
        # Update UI state
        self.root.after(0, self._update_ui_after_scraping, success)
    
    def _update_ui_after_scraping(self, success):
        """Update UI after scraping completion (runs in main thread)"""
        self.start_button.config(state=tk.NORMAL)
        self.stop_button.config(state=tk.DISABLED)
        self.progress_bar.stop()
        
        if success:
            self.update_status("✅ Scraping completed successfully!")
            messagebox.showinfo("Success", 
                              f"Scraping completed successfully!\n\n"
                              f"Files saved to: {self.output_dir_var.get()}\n"
                              f"Format: {self.format_var.get()}")
        else:
            self.update_status("❌ Scraping failed. Check logs for details.")
            messagebox.showerror("Error", 
                               "Scraping failed. Please check the logs for more details.")
    
    def stop_scraping(self):
        """Stop the scraping process"""
        if self.scraper:
            self.scraper.stop_scraping()
            self.update_status("Stopping scraper...")
    
    def open_output_folder(self):
        """Open the output folder in file explorer"""
        output_dir = Path(self.output_dir_var.get())
        
        if not output_dir.exists():
            messagebox.showwarning("Warning", f"Output directory does not exist: {output_dir}")
            return
            
        try:
            import subprocess
            import sys
            
            if sys.platform == "win32":
                subprocess.run(['explorer', str(output_dir)])
            elif sys.platform == "darwin":  # macOS
                subprocess.run(['open', str(output_dir)])
            else:  # Linux
                subprocess.run(['xdg-open', str(output_dir)])
                
        except Exception as e:
            messagebox.showerror("Error", f"Could not open folder: {e}")


def main():
    """Main function to run the GUI application"""
    root = tk.Tk()
    
    # Set application icon and styling
    try:
        # Try to set a nice theme if available
        style = ttk.Style()
        available_themes = style.theme_names()
        
        # Prefer modern themes
        preferred_themes = ['vista', 'xpnative', 'winnative', 'clam']
        for theme in preferred_themes:
            if theme in available_themes:
                style.theme_use(theme)
                break
                
    except Exception:
        pass  # Use default theme if styling fails
    
    # Center window on screen
    root.update_idletasks()
    width = root.winfo_width()
    height = root.winfo_height()
    x = (root.winfo_screenwidth() // 2) - (width // 2)
    y = (root.winfo_screenheight() // 2) - (height // 2)
    root.geometry(f"{width}x{height}+{x}+{y}")
    
    # Create and run application
    app = TableScraperGUI(root)
    
    # Handle window closing
    def on_closing():
        if app.scraper and app.scraper.is_running:
            if messagebox.askokcancel("Quit", "Scraping is in progress. Do you want to stop and quit?"):
                app.stop_scraping()
                root.after(1000, root.destroy)  # Give time for cleanup
        else:
            root.destroy()
    
    root.protocol("WM_DELETE_WINDOW", on_closing)
    
    # Start the GUI
    root.mainloop()


if __name__ == "__main__":
    main()