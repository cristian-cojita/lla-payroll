from dataclasses import dataclass


@dataclass
class PayrollInfo:
    """A data class to hold payroll date information."""
    execute_on_date: str
    payroll_from: str
    payroll_to: str
    weekNumber: str


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

DEFAULT_PAYROLL_METRICS_2026 = PayrollMetrics(
    paycor_id=2,
    sales=2,
    car_bonus=3,
    alignments=4,
    tire_units=5,
    fluids=6,
    brake_sales=7,
    clocked_hours=5,
    hours=8,
    overtime=9,
    labor_h=10,
    tech_h=11,
    invoiced_l=12,
    total_payroll=20,
    date=22,
)