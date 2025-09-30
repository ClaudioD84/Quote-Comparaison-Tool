from dataclasses import dataclass, field, asdict
from typing import List, Dict, Optional, Union

@dataclass
class ParsedOffer:
    """
    A standardized dataclass to hold all information extracted from a leasing offer.
    This structure ensures consistency after AI parsing from various document formats.
    """
    filename: str
    vendor: Optional[str] = None
    vehicle_description: Optional[str] = None
    
    # Contract terms: Differentiating between the maximum allowed and the actual offer
    max_duration_months: Optional[int] = None
    max_total_mileage: Optional[int] = None
    offer_duration_months: Optional[int] = None
    offer_total_mileage: Optional[int] = None
    
    # Core financial details
    monthly_rental: Optional[float] = None
    total_monthly_lease: Optional[float] = None # Often a more inclusive monthly cost
    currency: Optional[str] = None
    
    # One-time costs
    upfront_costs: Optional[float] = None
    deposit: Optional[float] = None
    admin_fees: Optional[float] = None
    
    # Mileage rates
    excess_mileage_rate: Optional[float] = None
    unused_mileage_rate: Optional[float] = None
    
    # Vehicle specifications
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
    battery_range: Optional[float] = None # Especially for EVs
    
    # Detailed cost breakdown (investment)
    vehicle_price: Optional[float] = None
    options_price: Optional[float] = None
    accessories_price: Optional[float] = None
    delivery_cost: Optional[float] = None
    registration_tax: Optional[float] = None
    total_net_investment: Optional[float] = None
    taxation_value: Optional[float] = None
    
    # Detailed cost breakdown (monthly services)
    financial_rate: Optional[float] = None
    depreciation_interest: Optional[float] = None
    maintenance_repair: Optional[float] = None
    insurance_cost: Optional[float] = None
    green_tax: Optional[float] = None
    management_fee: Optional[float] = None
    roadside_assistance: Optional[float] = None
    tyres_cost: Optional[float] = None
    maintenance_included: Optional[bool] = None

    # Customer and driver details
    driver_name: Optional[str] = None
    customer: Optional[str] = None
    
    # Itemized lists for options and accessories
    options_list: List[Dict[str, Union[str, float]]] = field(default_factory=list)
    accessories_list: List[Dict[str, Union[str, float]]] = field(default_factory=list)
    
    # Metadata from the parsing process
    parsing_confidence: float = 0.0
    warnings: List[str] = field(default_factory=list)


