
from datetime import date, datetime, time
from sqlalchemy import Date, DateTime, Integer, String, Time, Text
from sqlalchemy.orm import Mapped, mapped_column, relationship

from app.database import Base


class DailyEntry(Base):
    __tablename__ = "daily_entries"

    id: Mapped[int] = mapped_column(Integer, primary_key=True, index=True)
    user_id: Mapped[int] = mapped_column(Integer, index=True)
    entry_date: Mapped[date] = mapped_column(Date)
    site_id: Mapped[str] = mapped_column(String(100), index=True)
    area: Mapped[str | None] = mapped_column(String(100), nullable=True)
    stay_place: Mapped[str | None] = mapped_column(String(150), nullable=True)
    team_leader: Mapped[str | None] = mapped_column(String(150), nullable=True)
    work_start: Mapped[time | None] = mapped_column(Time, nullable=True)
    end_time: Mapped[time | None] = mapped_column(Time, nullable=True)
    on_site_duration: Mapped[str | None] = mapped_column(String(50), nullable=True)
    travel_stay_type: Mapped[str | None] = mapped_column(String(50), nullable=True)
    reservation_type: Mapped[str | None] = mapped_column(String(100), nullable=True)
    expense: Mapped[str | None] = mapped_column(String(100), nullable=True)
    team_members_text: Mapped[str | None] = mapped_column(Text, nullable=True)
    car_details: Mapped[str | None] = mapped_column(String(150), nullable=True)
    action_details: Mapped[str | None] = mapped_column(Text, nullable=True)
    comment: Mapped[str | None] = mapped_column(Text, nullable=True)
    submission_status: Mapped[str] = mapped_column(String(20), default="draft")
    manager_approval: Mapped[str] = mapped_column(String(20), default="n/a")
    created_at: Mapped[datetime] = mapped_column(DateTime, default=datetime.utcnow)
    updated_at: Mapped[datetime] = mapped_column(DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)
    evaluation = relationship("SiteEvaluation", uselist=False, primaryjoin="DailyEntry.id==foreign(SiteEvaluation.daily_entry_id)")
