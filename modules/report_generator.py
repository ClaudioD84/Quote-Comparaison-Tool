import pandas as pd
import io
import logging
from typing import List, Dict, Any
from xlsxwriter.utility import xl_rowcol_to_cell

# Import the data model and comparator logic
from modules.models import ParsedOffer, asdict
from modules.offer_comparator import OfferComparator, get_offer_diff, calculate_similarity_score

def _apply_excel_formatting(writer: pd.ExcelWriter, df: pd.DataFrame):
    """Applies all xlsxwriter formatting to the final report."""
    workbook = writer.book
    worksheet = writer.sheets['Quotation Comparison']
    
    # --- Define formats ---
    formats = {
        'header': workbook.add_format({'bold': True, 'bg_color': '#DDEBF7', 'border': 1, 'text_wrap': True, 'valign': 'top'}),
        'bold': workbook.add_format({'bold': True}),
        'winner_header': workbook.add_format({'bold': True, 'bg_color': '#C6EFCE', 'font_color': '#006100', 'border': 1}),
        'winner_cell': workbook.add_format({'bg_color': '#C6EFCE', 'font_color': '#006100'}),
        'wrap': workbook.add_format({'text_wrap': True, 'valign': 'top'}),
        'green': workbook.add_format({'bg_color': '#C6EFCE'}),
        'red': workbook.add_format({'bg_color': '#FFC7CE'}),
        'orange': workbook.add_format({'bg_color': '#FFEB9C'}),
        'percent': workbook.add_format({'num_format': '0.0"%"'}),
        'currency': workbook.add_format({'num_format': '#,##0.00'}),
    }

    # Apply header format
    for col_num, value in enumerate(df.columns.values):
        worksheet.write(0, col_num, value, formats['header'])

    # Find winner column index
    winner_row = df[df['Field'] == 'Winner'].values.flatten().tolist()
    winner_col_idx = winner_row.index("🥇 Winner") if "🥇 Winner" in winner_row else -1

    # --- Apply formatting row by row ---
    for r_idx, row in enumerate(df.itertuples(index=False), start=1):
        field_name = str(row[0])
        
        # Bold section headers
        if field_name in ['Investment', 'Taxation', 'Duration & Mileage', 'Financial rate', 'Service rate', 'Monthly fee', 'Excess / unused km', 'Equipment', 'Gap analysis', 'Cost Analysis']:
            worksheet.set_row(r_idx, None, formats['bold'])
        
        # Apply text wrapping for specific fields
        if field_name in ['Gap analysis', 'Additional equipment', 'Vehicle Description']:
            worksheet.set_row(r_idx, 60, formats['wrap'])
        
        # Apply percentage format
        if field_name == 'Vehicle description correspondence':
             for c_idx in range(1, len(row)):
                if isinstance(row[c_idx], (int, float)):
                    worksheet.write(r_idx, c_idx, row[c_idx], formats['percent'])

        # --- Conditional Formatting for Specifications ---
        spec_fields = ['Vehicle Description', 'Manufacturer', 'Model', 'Version', 'Internal colour', 'External colour', 'Fuel type']
        if field_name in spec_fields and len(row) > 2:
            ref_val = row[1]
            worksheet.write(r_idx, 1, ref_val, formats['green']) # Reference is always green
            for c_idx in range(2, len(row)):
                comp_val = row[c_idx]
                score = calculate_similarity_score(str(ref_val), str(comp_val))
                fmt = formats['green'] if score > 95 else (formats['orange'] if score > 80 else formats['red'])
                worksheet.write(r_idx, c_idx, comp_val, fmt)

        # Winner column highlighting
        if winner_col_idx != -1:
            worksheet.write(r_idx, winner_col_idx, row[winner_col_idx]) # Write data first
            worksheet.conditional_format(r_idx, winner_col_idx, r_idx, winner_col_idx, {'type': 'no_blanks', 'format': formats['winner_cell']})
            worksheet.write(0, winner_col_idx, df.columns[winner_col_idx], formats['winner_header'])


    # Set column widths
    worksheet.set_column('A:A', 40) # Field names
    worksheet.set_column('B:Z', 30) # Offer data columns

def generate_excel_report(offers: List[ParsedOffer]) -> io.BytesIO:
    """Orchestrates the generation of the final, formatted Excel report."""
    if len(offers) < 2:
        # Fallback for single offer, though UI should prevent this
        df = pd.DataFrame([asdict(o) for o in offers]).set_index('filename').transpose()
        buffer = io.BytesIO()
        df.to_excel(buffer, sheet_name='Quotation Details')
        return buffer

    ref_offer = offers[0]
    comp_offers = offers[1:]
    
    # --- Define the structure of the report ---
    template_fields = [
        'Quote number', 'Driver name', 'Vehicle Description', 'Manufacturer', 'Model',
        'Version', 'Internal colour', 'External colour', 'Fuel type', 'No. doors', 'HP', 
        'C02 emission WLTP (g/km)', 'Battery range', '', 'Equipment', 'Additional equipment',
        'Additional equipment price', '', 'Investment', 'Vehicle list price (excl. VAT, excl. options)',
        'Options (excl. taxes)', 'Accessories (excl. taxes)', 'Delivery cost', 'Registration tax',
        'Total net investment', '', 'Gap analysis', 'Vehicle description correspondence', '', 'Taxation', 
        'Taxation value', '', 'Duration & Mileage', 'Term (months)', 'Mileage per year (in km)', 
        '', 'Financial rate', 'Monthly financial rate (depreciation + interest)', '', 'Service rate',
        'Maintenance & repair', 'Insurance', 'Management fee', 'Tyres (summer and winter)', 
        'Road side assistance', '', 'Monthly fee', 'Total monthly lease ex. VAT', '', 
        'Excess / unused km', 'Excess kilometers', 'Unused kilometers'
    ]
    
    report_data = {'Field': template_fields}
    
    # --- Populate the report data for each offer ---
    for offer in offers:
        col_data = []
        offer_dict = offer.to_dict()
        # Mapping from template field to ParsedOffer attribute
        field_map = {
            'Quote number': 'quote_number', 'Driver name': 'driver_name', 'Vehicle Description': 'vehicle_description',
            'Manufacturer': 'manufacturer', 'Model': 'model', 'Version': 'version', 'Internal colour': 'internal_colour',
            'External colour': 'external_colour', 'Fuel type': 'fuel_type', 'No. doors': 'num_doors', 'HP': 'hp',
            'C02 emission WLTP (g/km)': 'c02_emission', 'Battery range': 'battery_range',
            'Vehicle list price (excl. VAT, excl. options)': 'vehicle_price', 'Options (excl. taxes)': 'options_price',
            'Accessories (excl. taxes)': 'accessories_price', 'Delivery cost': 'delivery_cost',
            'Registration tax': 'registration_tax', 'Total net investment': 'total_net_investment',
            'Taxation value': 'taxation_value', 'Term (months)': 'offer_duration_months',
            'Monthly financial rate (depreciation + interest)': 'depreciation_interest', 'Maintenance & repair': 'maintenance_repair',
            'Insurance': 'insurance_cost', 'Management fee': 'management_fee', 'Tyres (summer and winter)': 'tyres_cost',
            'Road side assistance': 'roadside_assistance', 'Total monthly lease ex. VAT': 'total_monthly_lease',
            'Excess kilometers': 'excess_mileage_rate', 'Unused kilometers': 'unused_mileage_rate'
        }
        
        for field in template_fields:
            if not field:
                col_data.append('')
                continue
            
            # Handle special computed/formatted fields
            if field == 'Additional equipment':
                equip_list = (offer_dict.get('options_list') or []) + (offer_dict.get('accessories_list') or [])
                names = [str(item.get('name', '')) for item in equip_list if isinstance(item, dict) and item.get('name')]
                col_data.append(", ".join(sorted(names)))
            elif field == 'Additional equipment price':
                 equip_list = (offer_dict.get('options_list') or []) + (offer_dict.get('accessories_list') or [])
                 total_price = sum(item.get('price') or 0 for item in equip_list if isinstance(item, dict))
                 col_data.append(total_price if total_price > 0 else (offer.options_price or 0) + (offer.accessories_price or 0))
            elif field == 'Mileage per year (in km)':
                if offer.offer_total_mileage and offer.offer_duration_months:
                    col_data.append(int(offer.offer_total_mileage / (offer.offer_duration_months / 12)))
                else:
                    col_data.append(None)
            elif field == 'Gap analysis':
                col_data.append(get_offer_diff(ref_offer, offer) if offer.filename != ref_offer.filename else 'N/A')
            elif field == 'Vehicle description correspondence':
                score = calculate_similarity_score(ref_offer.vehicle_description, offer.vehicle_description)
                col_data.append(score if offer.filename != ref_offer.filename else 100.0)
            elif field in field_map:
                col_data.append(offer_dict.get(field_map[field]))
            else:
                 col_data.append(None) # Default for unmapped fields
        
        report_data[offer.vendor or offer.filename] = col_data

    df = pd.DataFrame(report_data)
    
    # --- Add Cost Analysis Section at the end ---
    comp = OfferComparator(offers)
    cost_df = comp.generate_comparison_report()
    if not cost_df.empty:
        min_cost = cost_df['total_contract_cost'].min()
        
        # Align cost data with main df columns
        cost_data_aligned = cost_df.set_index('filename').reindex([o.filename for o in offers])
        
        df.loc[len(df)] = [''] * len(df.columns)
        df.loc[len(df)] = ['Cost Analysis'] + [''] * (len(df.columns) - 1)
        df.loc[len(df)] = ['Total Cost (excl. VAT)'] + cost_data_aligned['total_contract_cost'].tolist()
        df.loc[len(df)] = ['Monthly Cost (excl. VAT)'] + cost_data_aligned['cost_per_month'].tolist()
        df.loc[len(df)] = ['Winner'] + ["🥇 Winner" if cost == min_cost else "" for cost in cost_data_aligned['total_contract_cost']]
        
    # --- Generate the Excel file in memory ---
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name='Quotation Comparison', index=False)
        _apply_excel_formatting(writer, df)
        
    buffer.seek(0)
    return buffer

