# Ex. 1
# # import tkinter module
# import tkinter as tk

# # initialize Tk root widget
# root = tk.Tk()

# # Label widget - first parameter is the name of the parent window (root)
# w = tk.Label(root, text='Click the button below to select a file')

# # Pack method tells Tk to fit the size of the window to the given text
# w.pack()

# # Tkinter event loop
# root.mainloop()

# Ex. 2
import tkinter as tk

root = tk.Tk()
logo = tk.PhotoImage(file="python_logo_small.gif")

w1 = tk.Label(root, image=logo).pack(side="right")

explanation = """At present, only GIF and PPM/PGM
formats are supported, but an interface 
exists to allow additional image file
formats to be added easily."""

w2 = tk.Label(root, 
              justify=tk.LEFT,
              padx = 10, 
              text=explanation).pack(side="left")
root.mainloop()