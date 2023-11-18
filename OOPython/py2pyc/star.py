import turtle
import random

# Membuat tampilan Turtle
screen = turtle.Screen()
screen.bgcolor("black")

# Membuat objek Turtle
star = turtle.Turtle()
star.shape("triangle")
star.color("white")
star.speed(0)

# Fungsi untuk membuat bintang
def create_star():
    colors = ["red", "orange", "yellow", "green", "blue", "purple"]
    for _ in range(5):
        color = random.choice(colors)
        star.color(color)
        star.forward(100)
        star.right(144)

# Menggerakkan bintang secara acak
def move_star():
    x = random.randint(-200, 200)
    y = random.randint(-200, 200)
    star.penup()
    star.goto(x, y)
    star.pendown()

# Membuat bintang pertama
# Menggerakkan bintang secara terus-menerus
while True:
    create_star()
    move_star()

# Menutup jendela saat selesai
screen.mainloop()
