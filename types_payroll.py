from dataclasses import dataclass


@dataclass
class PayrollInfo:
    """A data class to hold payroll date information."""
    execute_on_date: str
    payroll_from: str
    payroll_to: str


@dataclass
class PayrollMetrics:
    """Data class for payroll metrics with their column numbers."""
    date: int
    sales: int
    car_bonus: int
    alignments: int
    tire_units: int
    fluids: int
    brake_sales: int
    hours: int
    overtime: int
    labor_h: int
    tech_h: int
    invoiced_l: int
    total_payroll: int
    paycor_id: int


# Default metric values
DEFAULT_METRICS = PayrollMetrics(
    paycor_id=2,
    sales=6,
    hours=7,
    overtime=8,
    car_bonus=8,
    labor_h=9,
    tech_h=10,
    alignments=10,
    invoiced_l=11,
    tire_units=12,
    fluids=14,
    brake_sales=16,
    total_payroll=16,
    date=18,
)