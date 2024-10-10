class Room:
    def __init__(self, no, type_r, price, availability):
        self.room_number = no
        self.room_type = type_r
        self.price = price
        self.is_available = True


class RoomManager:
    def __init__(self):
        self.rooms = {}

    def add_room(self, room_no, room_type, room_price):
        r1 = Room(room_no, room_type, room_price,True)
        self.rooms[room_no] = r1
        print("room was added")

    def check_availability(self, room_type):
        available_rooms = [room for room in self.rooms.values() if room.room_type == room_type and room.is_available]
        return available_rooms