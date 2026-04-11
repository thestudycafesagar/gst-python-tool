                                login_success = True; break
                        except: pass
                        
                        if "Invalid Password" in driver.page_source: return "Failed", "Invalid Password"
                        
                        try:
                            dual = driver.find_elements(By.XPATH, "//button[contains(text(), 'Login Here')]")
                            if dual and dual[0].is_displayed():
                                driver.execute_script("arguments[0].click();", dual[0])
                                time.sleep(2)
                        except: pass
                    if login_success: break
                except Exception as e:
                    self.log(f"   ⚠️ Login Error: {str(e)[:50]}")
                    if login_attempt < 3:
                        time.sleep(2)

            if not login_success: return "Failed", "Login Timeout"

            # ==========================================================
            # STEP A: Click "e-File" from the top navigation menu
            # ==========================================================
            self.log("   🔹 Step A: Clicking 'e-File' menu...")
            efile_clicked = False
            for attempt in range(3):
                try:
                    efile_btn = wait.until(EC.element_to_be_clickable((
                        By.XPATH,
                        "//span[contains(@class,'mdc-button__label') and normalize-space(text())='e-File']"
                    )))
                    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", efile_btn)
                    time.sleep(0.5)
                    driver.execute_script("arguments[0].click();", efile_btn)
                    self.log("   ✅ 'e-File' clicked.")
                    efile_clicked = True
                    break
                except Exception as e:
                    if attempt == 2:
                        self.log(f"   ❌ Failed to click 'e-File': {str(e)[:60]}")
                        return "Failed", "Could not click e-File"
                    self.log(f"   ⚠️ Retry {attempt+1}/3 for 'e-File'...")
                    time.sleep(1.5)

            # ==========================================================
            # STEP B: Click "e-Pay Tax" from the dropdown
            # ==========================================================
            self.log("   🔹 Step B: Clicking 'e-Pay Tax' menu item...")
            epay_clicked = False
            for attempt in range(3):
                try:
                    epay_item = wait.until(EC.element_to_be_clickable((
                        By.XPATH,
                        "//span[contains(@class,'mat-mdc-menu-item-text')]//span[normalize-space(text())='e-Pay Tax']"
                    )))
                    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", epay_item)
                    time.sleep(0.4)
                    driver.execute_script("arguments[0].click();", epay_item)
                    self.log("   ✅ 'e-Pay Tax' clicked.")
                    epay_clicked = True
                    break
                except Exception as e:
                    if attempt == 2:
                        self.log(f"   ❌ Failed to click 'e-Pay Tax': {str(e)[:60]}")
                        return "Failed", "Could not click e-Pay Tax"
                    self.log(f"   ⚠️ Retry {attempt+1}/3 for 'e-Pay Tax'...")
                    time.sleep(1.5)

            # ==========================================================
            # STEP B-2: Handle Applicable Income Tax Act Selection
            # ==========================================================
            try:
                time.sleep(2)
                # Check if the "Select Applicable Income Tax Act" text exists
                act_selection = driver.find_elements(By.XPATH, "//*[contains(text(), 'Select Applicable Income Tax Act')]")
                if act_selection:
                    self.log("   🔹 Step B-2: 'Income Tax Act' selection screen detected.")
                    
                    # Try to select the 'Income-tax Act, 1961' radio button (usually for older years)
                    # or 'Income-tax Act, 2025' depending on requirement. Defaulting to 1961 for older years logic.
                    try:
                        act_1961_radio = driver.find_element(By.XPATH, "//div[contains(text(), 'Income-tax Act, 1961')]/ancestor::label")
                        driver.execute_script("arguments[0].click();", act_1961_radio)
                        self.log("   ✅ Selected 'Income-tax Act, 1961'.")
                        time.sleep(0.5)
                    except:
                        self.log("   ⚠️ Could not specifically click 1961 act radio, proceeding with default.")
                    
                    # Click Continue button
                    try:
                        continue_btn = driver.find_element(By.XPATH, "//button[contains(., 'Continue')]")
                        driver.execute_script("arguments[0].click();", continue_btn)
                        self.log("   ✅ Clicked 'Continue' on Act Selection screen.")
                        time.sleep(2)
                    except:
                        self.log("   ⚠️ Could not click 'Continue' button.")
            except Exception as e:
                # Screen might not appear at all
                pass

            # ==========================================================
            # STEP C: Click "Payment History" tab
            # ==========================================================
            self.log("   🔹 Step C: Clicking 'Payment History' tab...")
            time.sleep(2)  # Allow the e-Pay Tax page to load
            for attempt in range(3):
                try:
                    pay_hist_tab = wait.until(EC.element_to_be_clickable((
                        By.XPATH,
                        "//span[contains(@class,'mdc-tab__text-label') and contains(normalize-space(text()),'Payment History')]"
                    )))
                    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", pay_hist_tab)
                    time.sleep(0.4)
                    driver.execute_script("arguments[0].click();", pay_hist_tab)
                    self.log("   ✅ 'Payment History' tab clicked.")
                    break
                except Exception as e:
                    if attempt == 2:
                        self.log(f"   ❌ Failed to click 'Payment History': {str(e)[:60]}")
                        return "Failed", "Could not click Payment History tab"
                    self.log(f"   ⚠️ Retry {attempt+1}/3 for 'Payment History'...")
                    time.sleep(1.5)

            # Wait for Payment History AG Grid table to load
            time.sleep(3)
            self.log("   ✅ Payment History page loaded.")

            # ==========================================================
            # STEP D: Fetch all Assessment Years from the AG Grid table
            # ==========================================================
            self.log("   🔹 Step D: Reading Assessment Years from table...")
            all_years = []
            for fetch_attempt in range(3):
                try:
                    # Wait for at least one assessmentYear cell to appear
                    wait.until(EC.presence_of_element_located((
                        By.CSS_SELECTOR, "[col-id='assessmentYear'].ag-cell-value"
                    )))
                    time.sleep(1)  # Let all rows render
                    year_cells = driver.find_elements(
                        By.CSS_SELECTOR, "[col-id='assessmentYear'].ag-cell-value"
                    )
                    raw_years = [c.text.strip() for c in year_cells if c.text.strip()]
                    # Deduplicate while preserving order
                    seen = set()
                    for y in raw_years:
                        if y and y not in seen:
                            seen.add(y)
                            all_years.append(y)
                    if all_years:
                        break
                    time.sleep(1.5)
                except Exception as e:
                    if fetch_attempt == 2:
                        self.log(f"   ❌ Could not read year table: {str(e)[:60]}")
                        return "Failed", "Could not read Assessment Year table"
                    self.log(f"   ⚠️ Retry {fetch_attempt+1}/3 reading year table...")
                    time.sleep(2)

            if not all_years:
                self.log("   ⚠️ No challan records found in Payment History.")
                return "Success", "No Challan Records Found"

            # Sort years descending (e.g. ["2026-27","2025-26","2024-25"])
            def year_sort_key(y):
                try: return int(y.split("-")[0])
                except: return 0
            all_years.sort(key=year_sort_key, reverse=True)
            self.log(f"   📋 All Years Found: {', '.join(all_years)}")

            # ==========================================================
            # STEP E: Filter years based on selected Download Filter
            # ==========================================================
            if self.year_mode == "Current Year":
                selected_years = all_years[:1]
            elif self.year_mode == "Last 2 Years":
                selected_years = all_years[:2]
            else:  # "All History"
                selected_years = all_years

            self.log(f"   🎯 Selected Years to Download ({self.year_mode}): {', '.join(selected_years)}")

            # ==========================================================
            # STEP F: For each selected year — click ⋮ (more_vert) → Download
            # Angular Material menus render in a global overlay (outside the row),
            # so we must locate the row first, click its icon button, then find
            # the Download item in the CDK overlay container.
            # ==========================================================
            downloaded_years = []
            failed_years = []

            for year in selected_years:
                self.log(f"   🔹 Processing Year: {year}")

                for dl_attempt in range(3):
                    try:
                        # Dismiss any leftover open menu
                        try:
                            driver.execute_script("document.body.click();")
                            time.sleep(0.5)
                        except: pass

                        # --- F1: Find the AG Grid row whose assessmentYear cell = year ---
                        year_cell_xpath = (
                            f"//div[@col-id='assessmentYear' and "
                            f"contains(@class,'ag-cell') and "
                            f"normalize-space(text())='{year}']"
                        )
                        year_cell = wait.until(EC.presence_of_element_located((By.XPATH, year_cell_xpath)))
                        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", year_cell)
                        time.sleep(0.5)

                        # --- F2: From that cell climb to role="row", find the more_vert button ---
                        # The action button contains a <mat-icon> with text "more_vert"
                        row_el = year_cell.find_element(By.XPATH, "./ancestor::div[@role='row'][1]")
                        more_btn = row_el.find_element(
                            By.XPATH,
                            ".//button[contains(@class,'mat-mdc-icon-button')]"
                            "[.//mat-icon[normalize-space(text())='more_vert']]"
                        )
                        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", more_btn)
                        time.sleep(0.3)
                        driver.execute_script("arguments[0].click();", more_btn)
                        self.log(f"      ✅ ⋮ menu opened for {year}.")

                        # --- F3: Wait for Angular Material overlay, then click Download ---
                        # Mat menus render inside .cdk-overlay-container at the document level
                        download_btn = WebDriverWait(driver, 10).until(
                            EC.element_to_be_clickable((
                                By.XPATH,
                                "//div[contains(@class,'cdk-overlay-container')]"
                                "//button[contains(@class,'mat-mdc-menu-item')]"
                                "[.//span[normalize-space(text())='Download']]"
                            ))
                        )
                        driver.execute_script("arguments[0].click();", download_btn)
                        self.log(f"      ✅ 'Download' clicked for {year}.")
                        time.sleep(4)  # Wait for file download to initiate

                        downloaded_years.append(year)
                        break

                    except Exception as e:
                        # Dismiss menu before retry
                        try: driver.execute_script("document.body.click();")
                        except: pass
                        time.sleep(1)

                        if dl_attempt == 2:
                            self.log(f"      ❌ Download failed for {year}: {str(e)[:80]}")
                            failed_years.append(year)
                        else:
                            self.log(f"      ⚠️ Retry {dl_attempt+1}/3 for year {year}...")
                            time.sleep(2)

            summary = f"Downloaded: {', '.join(downloaded_years) or 'None'}"
