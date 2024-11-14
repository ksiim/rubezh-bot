import asyncio

from models.databases import Session
from models.dbs.models import *

from sqlalchemy import insert, inspect, or_, select, text, delete


class Orm:
    
    @staticmethod
    async def delete_last_messages():
        async with Session() as session:
            query = (
                select(Message_)
            )
            messages = (await session.execute(query)).scalars().all()
            message_ids = [message.message_id for message in messages]
            for message in messages:
                await session.delete(message)
            await session.commit()
            return message_ids

    @staticmethod
    async def add_message(message, is_head=False):
        async with Session() as session:
            message = Message_(
                message_id=message.message_id,
                is_head=is_head
            )
            session.add(message)
            await session.commit()

    @staticmethod
    async def create_user(message):
        if await Orm.get_user_by_telegram_id(message.from_user.id) is None:
            async with Session() as session:
                user = User(
                    full_name=message.from_user.full_name,
                    telegram_id=message.from_user.id,
                    username=message.from_user.username
                )
                session.add(user)
                await session.commit()

    @staticmethod
    async def get_user_by_telegram_id(telegram_id):
        async with Session() as session:
            query = select(User).where(User.telegram_id == telegram_id)
            user = (await session.execute(query)).scalar_one_or_none()
            return user

    @staticmethod
    async def get_all_users():
        async with Session() as session:
            query = select(User)
            users = (await session.execute(query)).scalars().all()
            return users
