from datetime import date, datetime
from sqlalchemy import Date, DateTime, Float, Integer, String, Text
from sqlalchemy.orm import Mapped, mapped_column

from app.database import Base


class SiteEvaluation(Base):
    __tablename__ = "site_evaluations"

    id: Mapped[int] = mapped_column(Integer, primary_key=True, index=True)
    daily_entry_id: Mapped[int] = mapped_column(Integer, index=True)
    user_id: Mapped[int] = mapped_column(Integer, index=True)
    eval_date: Mapped[date] = mapped_column(Date)
    site_id: Mapped[str] = mapped_column(String(100), index=True)
    area: Mapped[str | None] = mapped_column(String(100), nullable=True)
    stay_place: Mapped[str | None] = mapped_column(String(150), nullable=True)
    site_type: Mapped[str | None] = mapped_column(String(50), nullable=True)
    activity: Mapped[str | None] = mapped_column(Text, nullable=True)
    team_members_qty: Mapped[int | None] = mapped_column(Integer, nullable=True)
    tower_type: Mapped[str | None] = mapped_column(String(100), nullable=True)
    hu_rru_qty: Mapped[float | None] = mapped_column(Float, nullable=True)
    hu_rru_score: Mapped[float | None] = mapped_column(Float, nullable=True)
    er_rru_qty: Mapped[float | None] = mapped_column(Float, nullable=True)
    er_rru_score: Mapped[float | None] = mapped_column(Float, nullable=True)
    fiber_qty: Mapped[float | None] = mapped_column(Float, nullable=True)
    fiber_score: Mapped[float | None] = mapped_column(Float, nullable=True)
    dc_qty: Mapped[float | None] = mapped_column(Float, nullable=True)
    dc_score: Mapped[float | None] = mapped_column(Float, nullable=True)
    indoor_hu_survey_psu_qty: Mapped[float | None] = mapped_column(Float, nullable=True)
    indoor_hu_survey_psu_score: Mapped[float | None] = mapped_column(Float, nullable=True)
    sw_er_qty: Mapped[float | None] = mapped_column(Float, nullable=True)
    sw_er_score: Mapped[float | None] = mapped_column(Float, nullable=True)
    power_cabinet: Mapped[str | None] = mapped_column(String(10), nullable=True)
    power_cabinet_score: Mapped[float | None] = mapped_column(Float, nullable=True)
    batteries_addition: Mapped[str | None] = mapped_column(String(10), nullable=True)
    batteries_addition_score: Mapped[float | None] = mapped_column(Float, nullable=True)
    violations: Mapped[float | None] = mapped_column(Float, nullable=True)
    acceptance: Mapped[str | None] = mapped_column(String(200), nullable=True)
    performance: Mapped[float | None] = mapped_column(Float, nullable=True)
    final_score: Mapped[float | None] = mapped_column(Float, nullable=True)
    comment: Mapped[str | None] = mapped_column(Text, nullable=True)
    created_at: Mapped[datetime] = mapped_column(DateTime, default=datetime.utcnow)
    updated_at: Mapped[datetime] = mapped_column(DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)
