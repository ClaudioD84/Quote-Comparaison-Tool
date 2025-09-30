
import pandas as pd
import re
import difflib
from typing import List, Dict, Any, Tuple, Optional

# Import the data model from the models module
from .models import ParsedOffer

# A mapping to standardize common European currency symbols and codes
CURRENCY_MAP = {
    '€': 'EUR', 'eur': 'EUR', 'euro': 'EUR',
    '£': 'GBP', 'gbp': 'GBP',
    'kr.': 'DKK', 'kr': 'DKK', 'dkk': 'DKK',
    'chf': 'CHF', 'sek': 'SEK', 'nok': 'NOK',
    'pln': 'PLN', 'huf': 'HUF', 'czk': 'CZK',
}

def normalize_currency(currency_str: Optional[str]) -> Optional[str]:
    """Normalizes a currency string to a standard 3-letter code."""
    if not currency_str:
        return None
    return CURRENCY_MAP.get(currency_str.lower(), currency_str.upper())

def calculate_similarity_score(s1: str, s2: str) -> float:
    """Calculates a similarity score between two strings, ignoring common noise."""
    def preprocess(text: str) -> str:
        text = text.lower()
        text = re.sub(r'[^a-z0-9\s]', '', text) # Remove punctuation
        # Remove common, non-descriptive terms
        common_words = {'el', 'km', 'hp', 'auto', 'color', 'vehicle'}
        tokens = [word for word in text.split() if word not in common_words]
        return " ".join(tokens)

    s1_clean = preprocess(str(s1 or ''))
    s2_clean = preprocess(str(s2 or ''))
    
    # Use SequenceMatcher for a robust comparison
    matcher = difflib.SequenceMatcher(None, s1_clean, s2_clean)
    return matcher.ratio() * 100

def get_offer_diff(ref_offer: ParsedOffer, comp_offer: ParsedOffer) -> str:
    """
    Compares two offers and generates a human-readable summary of the key differences.
    """
    diffs = []
    
    # --- Compare key specification fields ---
    fields_to_compare = [
        'vehicle_description', 'manufacturer', 'model', 'version',
        'internal_colour', 'external_colour', 'fuel_type'
    ]
    for field in fields_to_compare:
        val1 = str(getattr(ref_offer, field, None) or "N/A").strip()
        val2 = str(getattr(comp_offer, field, None) or "N/A").strip()
        # Only report a difference if values are not similar
        if calculate_similarity_score(val1, val2) < 95.0:
            diffs.append(f"• {field.replace('_', ' ').title()}: '{val1}' vs '{val2}'")

    # --- Compare equipment lists ---
    equip1 = {item['name'].strip().lower() for item in ref_offer.options_list + ref_offer.accessories_list}
    equip2 = {item['name'].strip().lower() for item in comp_offer.options_list + comp_offer.accessories_list}
    
    if added := sorted(list(equip2 - equip1)):
        diffs.append(f"• Equipment Added: {', '.join(added).title()}")
    if removed := sorted(list(equip1 - equip2)):
        diffs.append(f"• Equipment Removed: {', '.join(removed).title()}")
        
    return "\n".join(diffs) if diffs else "No significant differences found in specifications."


class OfferComparator:
    """Handles the validation, cost calculation, and reporting for a list of offers."""

    def __init__(self, offers: List[ParsedOffer]):
        if not offers:
            raise ValueError("OfferComparator requires at least one offer to be initialized.")
        self.offers = offers

    def validate_offers(self) -> Tuple[bool, List[str]]:
        """
        Validates that a list of offers is suitable for a direct comparison.
        Checks for consistency in currency, duration, and mileage.
        """
        errors = []
        if len(self.offers) < 2:
            errors.append("At least two offers are needed for a comparison.")
            return False, errors

        # Check for consistent currency
        currencies = {normalize_currency(o.currency) for o in self.offers if o.currency}
        if len(currencies) > 1:
            errors.append(f"Mixed currencies detected: {', '.join(filter(None, currencies))}")

        # Check for consistent duration
        durations = {o.offer_duration_months for o in self.offers if o.offer_duration_months}
        if len(durations) > 1:
            errors.append(f"Mismatched contract durations: {', '.join(map(str, durations))} months")

        # Check for consistent mileage
        mileages = {o.offer_total_mileage for o in self.offers if o.offer_total_mileage}
        if len(mileages) > 1:
            errors.append(f"Mismatched contract mileages: {', '.join(map(str, mileages))} km")

        return not errors, errors

    def calculate_total_costs(self) -> List[Dict[str, Any]]:
        """
        Calculates the total cost over the contract term for each offer.
        """
        results = []
        for offer in self.offers:
            # Prioritize 'total_monthly_lease' if available, otherwise use 'monthly_rental'
            monthly_rate = offer.total_monthly_lease if offer.total_monthly_lease is not None else offer.monthly_rental
            
            # Ensure essential data for calculation is present
            if offer.offer_duration_months is None or monthly_rate is None:
                results.append({'vendor': offer.vendor, 'filename': offer.filename, 'error': 'Missing duration or monthly rate'})
                continue

            # Calculate total cost components
            total_lease_cost = monthly_rate * offer.offer_duration_months
            total_upfront_cost = (offer.upfront_costs or 0) + (offer.deposit or 0) + (offer.admin_fees or 0)
            total_cost = total_lease_cost + total_upfront_cost
            
            results.append({
                'vendor': offer.vendor,
                'filename': offer.filename,
                'duration_months': offer.offer_duration_months,
                'total_mileage': offer.offer_total_mileage,
                'monthly_rental': monthly_rate,
                'total_contract_cost': total_cost,
                'cost_per_month': total_cost / offer.offer_duration_months,
                'cost_per_km': total_cost / offer.offer_total_mileage if offer.offer_total_mileage else None,
                'currency': offer.currency
            })
        
        # Sort results by the lowest total cost
        return sorted(results, key=lambda x: x.get('total_contract_cost', float('inf')))

    def generate_comparison_report(self) -> pd.DataFrame:
        """
        Generates a pandas DataFrame summarizing the cost comparison.
        """
        cost_data = self.calculate_total_costs()
        if not cost_data or all('error' in d for d in cost_data):
            return pd.DataFrame()

        df = pd.DataFrame(cost_data)
        
        # Add a ranking column based on the total cost
        df['rank'] = df['total_contract_cost'].rank(method='min').astype(int)
        
        return df

