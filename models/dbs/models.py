from sqlalchemy import ForeignKey
from sqlalchemy.orm import relationship, Mapped, mapped_column
from models.databases import Base


class User(Base):
    __tablename__ = 'users'

    id: Mapped[int] = mapped_column(primary_key=True)
    telegram_id: Mapped[int] = mapped_column(unique=True)
    full_name: Mapped[str]
    username: Mapped[str] = mapped_column(nullable=True)
    admin: Mapped[bool] = mapped_column(default=False)


class Message_(Base):
    __tablename__ = 'messages'

    id: Mapped[int] = mapped_column(primary_key=True)
    message_id: Mapped[int] = mapped_column(unique=True)
    is_head: Mapped[bool] = mapped_column(default=False)
