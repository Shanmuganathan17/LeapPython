import Room

room_manager = Room.RoomManager()
room_manager.add_room(101, 'single', 1000)
room_manager.add_room(111, 'double', 2000)


print(room_manager)
print(room_manager.rooms.values())