from dataclasses import dataclass, field, asdict
from typing import List, Dict, Any, Optional, Union

# This helper function is used to generate a clean dictionary representation of the dataclass.
# It helps in creating the JSON schema description for the AI prompt and for displaying data.
def to_dict_helper(obj):
    return asdict(obj)

@dataclass
class ParsedOffer:
    """Standardized structure for parsed leasing offer data"""
    filename: str
    vendor: Optional[str] = None
    vehicle_description: Optional[str] = None
    max_duration_months: Optional[int] = None
    max_total_mileage: Optional[int] = None
    offer_duration_months: Optional[int] = None
    offer_total_mileage: Optional[int] = None
    monthly_rental: Optional[float] = None
    total_monthly_lease: Optional[float] = None
    currency: Optional[str] = None
    upfront_costs: Optional[float] = None
    deposit: Optional[float] = None
    admin_fees: Optional[float] = None
    excess_mileage_rate: Optional[float] = None
    unused_mileage_rate: Optional[float] = None
    quote_number: Optional[str] = None
    manufacturer: Optional[str] = None
    model: Optional[str] = None
    version: Optional[str] = None
    internal_colour: Optional[str] = None
    external_colour: Optional[str] = None
    fuel_type: Optional[str] = None
    num_doors: Optional[int] = None
    hp: Optional[int] = None
    c02_emission: Optional[float] = None
    battery_range: Optional[float] = None
    vehicle_price: Optional[float] = None
    options_price: Optional[float] = None
    accessories_price: Optional[float] = None
    delivery_cost: Optional[float] = None
    registration_tax: Optional[float] = None
    total_net_investment: Optional[float] = None
    taxation_value: Optional[float] = None
    financial_rate: Optional[float] = None
    depreciation_interest: Optional[float] = None
    maintenance_repair: Optional[float] = None
    insurance_cost: Optional[float] = None
    green_tax: Optional[float] = None
    management_fee: Optional[float] = None
    roadside_assistance: Optional[float] = None
    tyres_cost: Optional[float] = None
    maintenance_included: Optional[bool] = None
    driver_name: Optional[str] = None
    customer: Optional[str] = None
    options_list: List[Dict[str, Union[str, float]]] = field(default_factory=list)
    accessories_list: List[Dict[str, Union[str, float]]] = field(default_factory=list)
    parsing_confidence: float = 0.0
    warnings: List[str] = field(default_factory=list)

    def to_dict(self) -> Dict[str, Any]:
        """
        Provides a method to convert the dataclass instance to a dictionary,
        which is needed for displaying with st.json().
        """
        return to_dict_helper(self)

