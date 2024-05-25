extends Node

const CONFIG_FILE = "user://settings.cfg"

func _ready():
	load_window_settings()

func _notification(what):
	if what == NOTIFICATION_WM_CLOSE_REQUEST:
		print("Hello1")
		save_window_settings()

func save_window_settings():
	var config = ConfigFile.new()
	
	var window = get_window()
	
	config.set_value("window", "size", window.size)
	config.set_value("window", "position", window.position)
	
	var error = config.save(CONFIG_FILE)
	if error != OK:
		print("Failed to save window settings: ", error)

func load_window_settings():
	var config = ConfigFile.new()
	var error = config.load(CONFIG_FILE)
	
	if error == OK:
		var size = config.get_value("window", "size", Vector2(800, 600))  # Default size if not set
		var position = config.get_value("window", "position", Vector2(100, 100))  # Default position if not set
		var window = get_window()
		window.size = size
		window.position = position
	else:
		print("No settings file found or other error: ", error)

