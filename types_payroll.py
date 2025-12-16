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
    clocked_hours: int
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
DEFAULT_PAYROLL_METRICS = PayrollMetrics(
    paycor_id=2,
    clocked_hours=5,
    sales=7,
    hours=8,
    overtime=9,
    car_bonus=9,
    labor_h=10,
    tech_h=11,
    alignments=11,
    invoiced_l=12,
    tire_units=13,
    fluids=15,
    brake_sales=17,
    total_payroll=19,
    date=21,
)