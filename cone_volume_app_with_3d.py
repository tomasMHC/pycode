
import streamlit as st
import numpy as np
import matplotlib.pyplot as plt
from mpl_toolkits.mplot3d import Axes3D

# Function to calculate volume of liquid in an inverted truncated cone
def calculate_volume(R, r, H, h):
    if h > H:
        h = H
    r_h = r + (R - r) * h / H
    volume = (1/3) * np.pi * h * (r**2 + r * r_h + r_h**2)
    return volume

# Function to plot the 3D cone and water level
def plot_cone(R, r, H, h):
    fig = plt.figure(figsize=(8, 6))
    ax = fig.add_subplot(111, projection='3d')

    # Create mesh for the full cone
    z = np.linspace(0, H, 50)
    theta = np.linspace(0, 2 * np.pi, 50)
    Z, Theta = np.meshgrid(z, theta)
    R_mesh = r + (R - r) * Z / H
    X = R_mesh * np.cos(Theta)
    Y = R_mesh * np.sin(Theta)

    ax.plot_surface(X, Y, Z, alpha=0.3, color='gray', edgecolor='k')

    # Create mesh for the water level
    z_water = np.linspace(0, h, 30)
    Z_w, Theta_w = np.meshgrid(z_water, theta)
    R_w = r + (R - r) * Z_w / H
    X_w = R_w * np.cos(Theta_w)
    Y_w = R_w * np.sin(Theta_w)

    ax.plot_surface(X_w, Y_w, Z_w, alpha=0.6, color='blue', edgecolor='k')

    ax.set_xlabel('X')
    ax.set_ylabel('Y')
    ax.set_zlabel('Height')
    ax.set_title('Inverted Truncated Cone with Water Level')
    ax.set_box_aspect([1,1,H/10])  # Aspect ratio

    return fig

# Streamlit UI
st.title("Inverted Truncated Cone Volume Calculator")

R = st.slider("Top Radius (R)", min_value=1.0, max_value=5.0, value=2.0)
r = st.slider("Bottom Radius (r)", min_value=0.1, max_value=R, value=0.15)
H = st.slider("Total Height (H)", min_value=1.0, max_value=5.0, value=2.0)
h = st.slider("Water Height (h)", min_value=0.0, max_value=H, value=1.0)

volume = calculate_volume(R, r, H, h)
st.write(f"### Volume of liquid at height {h} is: {volume:.2f} cubic units")

fig = plot_cone(R, r, H, h)
st.pyplot(fig)
