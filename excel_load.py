import pandas as pd
import xlwings as xw

class Load:
    def __init__(self, transformer, workbook: xw.Book):
        self.transformer = transformer
        self.workbook = workbook
        self.molecular_weight = 12.01057

    @staticmethod
    def is_out_of_bounds(value, check_type):
        try:
            val = float(value)
        except (ValueError, TypeError):
            return False
        
        if check_type == 'QC_R':
            return val < 90 or val > 110
        elif check_type == 'MDL_R':
            return val < 45 or val > 145
        elif check_type == 'RPD':
            return val > 10
        else:
            return False

    def sample_groups(self):
        cleaned = self.transformer.clean_data()
        df = cleaned.copy()
        df["Sample ID"] = df["Sample ID"].astype("string").str.strip()
        qc_pattern = r"(?i)^(MDL|ICV|ICB|CCV\d+|CCB\d+|Rinse)$"
        samples_only = df[~df["Sample ID"].str.match(qc_pattern, na=False)]
        ordered_ids = self.get_unique_ordered_ids(samples_only)
        groups = self.build_sample_groups(samples_only, ordered_ids)
        return samples_only, groups

    def get_unique_ordered_ids(self, df):
        ordered_ids = []
        seen = set()
        for sid in df["Sample ID"]:
            if sid not in seen:
                seen.add(sid)
                ordered_ids.append(sid)
        return ordered_ids

    def build_sample_groups(self, df, ordered_ids):
        groups = []
        for sample_id in ordered_ids:
            group_df = df[df["Sample ID"] == sample_id]
            groups.append((sample_id, group_df))
        return groups

    def format_qc(self):
        df = self.transformer.df
        samples = df[df["Sample Type"] == "Samples"].copy()
        samples["Sample ID"] = samples["Sample ID"].astype("string").str.strip()

        qc_mask = samples["Sample ID"].str.match(r"(?i)^(MDL|ICV|CCV\d+)$", na=False)
        qcb_mask = samples["Sample ID"].str.match(r"(?i)^(ICB|CCB\d+)$", na=False)
        qc_samples = samples[qc_mask]
        qcb_samples = samples[qcb_mask]

        columns = ["Sample ID", "PPM C", "Mean ppm C", "%R", "%RPD", "Bounds"]
        records = []

        records.extend(self.build_qc_records(qc_samples, add_bounds_once=True))
        records.append({col: None for col in columns})
        records.extend(self.build_qcb_records(qcb_samples))
        records.append(self.build_qcb_average(qcb_samples))

        return pd.DataFrame(records, columns=columns)

    def build_qc_records(self, qc_samples, add_bounds_once=False):
        qc_targets = {"MDL": 0.2, "ICV": 18.0}
        records = []
        bounds_added = False

        for sample_id in qc_samples["Sample ID"].unique():
            group_df = qc_samples[qc_samples["Sample ID"] == sample_id]
            group_records = []
            sample_id_upper = str(sample_id).upper()

            for _, row in group_df.iterrows():
                group_records.append({
                    "Sample ID": row["Sample ID"],
                    "PPM C": row.get("PPM"),
                    "Mean ppm C": None,
                    "%R": None,
                    "%RPD": None,
                    "Bounds": None
                })
            mean_ppm = self.transformer.calculate_mean_ppm(group_df)
            target = 10.0 if sample_id_upper.startswith("CCV") else qc_targets.get(sample_id_upper)
            percent_r = self.transformer.calculate_percent_R(group_df, target_override=target)
            rpd = self.transformer.calculate_rpd(group_df, mean_ppm)

            # Add summary values to last row
            last_record = group_records[-1]
            last_record["Mean ppm C"] = mean_ppm
            last_record["%R"] = percent_r
            last_record["%RPD"] = rpd

            # Add bounds only to the very first data row
            if add_bounds_once and not bounds_added:
                group_records[0]["Bounds"] = "MDL %R: 45-145%, ICV/CCV %R: 90-110%"
                bounds_added = True
            else:
                last_record["Bounds"] = None

            records.extend(group_records)
        return records

    def build_qcb_records(self, qcb_samples):
        records = []
        for _, row in qcb_samples.iterrows():
            records.append({"Sample ID": row["Sample ID"], "PPM C": row.get("PPM"), "Mean ppm C": None, "%R": None, "%RPD": None, "Bounds": None})
        return records

    def build_qcb_average(self, qcb_samples):
        average_ppm = self.transformer.calculate_mean_ppm(qcb_samples)
        return {"Sample ID": "Average", "PPM C": average_ppm, "Mean ppm C": None, "%R": None, "%RPD": None, "Bounds": None}

    def format_samples(self):
        samples_only, groups = self.sample_groups()
        columns = ["Sample ID", "PPM C", "Mean ppm C", "%RPD", "umol/L C", "Bounds"]
        records = []
        bounds_added = False

        for sample_id, group_df in groups:
            group_records = self.build_sample_group_records(group_df)
            self.add_summary_to_last_record(group_df, group_records)

            # Add bounds only to the very first data row
            if not bounds_added:
                group_records[0]["Bounds"] = "RPD: ≤10%"
                bounds_added = True

            records.extend(group_records)
        return pd.DataFrame(records, columns=columns)

    def build_sample_group_records(self, group_df):
        group_records = []
        for _, row in group_df.iterrows():
            group_records.append({"Sample ID": row["Sample ID"], "PPM C": row.get("PPM"), "Mean ppm C": None, "%RPD": None, "umol/L C": None, "Bounds": None})
        return group_records

    def add_summary_to_last_record(self, group_df, group_records):
        mean_ppm = self.transformer.calculate_mean_ppm(group_df)
        rpd = self.transformer.calculate_rpd(group_df, mean_ppm)
        mean_umol = self.transformer.convert_to_umol_per_L(mean_ppm, self.molecular_weight)
        last_record = group_records[-1]
        last_record["Mean ppm C"] = mean_ppm
        last_record["%RPD"] = rpd
        last_record["umol/L C"] = mean_umol
        last_record["Bounds"] = None

    def format_reported_results(self):
        _, groups = self.sample_groups()
        records = []
        for sample_id, group_df in groups:
            mean_ppm = self.transformer.calculate_mean_ppm(group_df)
            umol = self.transformer.convert_to_umol_per_L(mean_ppm, self.molecular_weight)
            records.append({"Sample ID": sample_id, "umol/L C": umol})
        return pd.DataFrame(records, columns=["Sample ID", "umol/L C"])

    def export_all(self):
        try:
            self.write_sheets()
            self.apply_formatting()
            self.cleanup_xlwings_config()
        except Exception as e:
            # If anything fails, show the specific error
            xw.apps.active.alert(f"Error during export: {type(e).__name__}: {str(e)}", "ETL Pipeline Error")

    def cleanup_xlwings_config(self):
        """Remove the _xlwings.conf sheet if it exists"""
        try:
            if '_xlwings.conf' in [sheet.name for sheet in self.workbook.sheets]:
                self.workbook.sheets['_xlwings.conf'].delete()
        except Exception:
            pass  # Ignore if deletion fails

    def write_sheets(self):
        qc_df = self.format_qc()
        samples_df = self.format_samples()
        results_df = self.format_reported_results()

        sheets_to_write = {
            "QC": qc_df,
            "Samples": samples_df,
            "Reported Results": results_df,
        }

        # Get the last sheet to add new sheets after it (to the right)
        last_sheet = self.workbook.sheets[-1]

        for sheet_name, df in sheets_to_write.items():
            if sheet_name in [sheet.name for sheet in self.workbook.sheets]:
                ws = self.workbook.sheets[sheet_name]
                ws.clear_contents()
            else:
                # Add new sheet after the last sheet (to the right)
                ws = self.workbook.sheets.add(sheet_name, after=last_sheet)
                last_sheet = ws  # Update last_sheet so next one goes after this

            # Write DataFrame without index - use options to exclude index column
            ws.range('A1').options(index=False).value = df

    def apply_formatting(self):
        # Windows Excel COM API expects BGR integer format, not RGB tuple
        # Red in BGR: Blue=0, Green=0, Red=255 -> 0x0000FF = 255
        red_font_color = 255  # BGR format as integer
        self.format_qc_sheet(red_font_color)
        self.format_samples_sheet(red_font_color)

    def format_qc_sheet(self, red_font_color):
        try:
            ws = self.workbook.sheets['QC']
            used_range = ws.used_range
            max_row = used_range.last_cell.row

            # Find the %R column by reading the header row
            header_row = ws.range('1:1').value
            r_col_idx = None
            for idx, header in enumerate(header_row, start=1):
                if header == '%R':
                    r_col_idx = idx
                    break

            if r_col_idx is None:
                return  # %R column not found

            for row_idx in range(2, max_row + 1):
                try:
                    sample_id_cell = ws.range((row_idx, 1))
                    sample_id = sample_id_cell.value
                    r_cell = ws.range((row_idx, r_col_idx))
                    if r_cell.value is not None:
                        # Determine check type based on sample ID
                        if sample_id and 'MDL' in str(sample_id).upper():
                            check_type = 'MDL_R'
                        else:
                            check_type = 'QC_R'

                        if self.is_out_of_bounds(r_cell.value, check_type):
                            # Apply red font color using xlwings font property
                            r_cell.font.color = (255, 0, 0)
                except Exception:
                    continue  # Skip this row if there's an error
        except Exception:
            pass  # If sheet formatting fails completely, continue

    def format_samples_sheet(self, red_font_color):
        try:
            ws = self.workbook.sheets['Samples']
            used_range = ws.used_range
            max_row = used_range.last_cell.row

            # Find the %RPD column by reading the header row
            header_row = ws.range('1:1').value
            rpd_col_idx = None
            for idx, header in enumerate(header_row, start=1):
                if header == '%RPD':
                    rpd_col_idx = idx
                    break

            if rpd_col_idx is None:
                return  # %RPD column not found

            for row_idx in range(2, max_row + 1):
                try:
                    rpd_cell = ws.range((row_idx, rpd_col_idx))
                    if rpd_cell.value is not None:
                        if self.is_out_of_bounds(rpd_cell.value, 'RPD'):
                            # Apply red font color using xlwings font property
                            rpd_cell.font.color = (255, 0, 0)
                except Exception:
                    continue  # Skip this row if there's an error
        except Exception:
            pass  # If sheet formatting fails completely, continue