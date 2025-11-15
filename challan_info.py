from dataclasses import dataclass

@dataclass
class ChallanInfo:
    file_name: str = None
    tender_date: str = None
    challan_no: str = None
    bsr_code: str = None
    section_raw: str = None
    section: str = None
    financial_year: str = None
    total_amount: int = 0
    tax_amount: int = 0
    interest_amount: int = 0
    other_amount: int = 0
    firm_name: str = None

    # Filler entries
    book_entry: str = "No"
    fee_amount: int = 0
    sa_entry: str = "Self Assessment-200"
