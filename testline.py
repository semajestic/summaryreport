import matplotlib.pyplot as plt

# Create a figure and axis
fig, ax = plt.subplots(figsize=(8.5, 11))

# Set the signatories' names
names = ['John Doe', 'Jane Smith']

# Set the y-coordinate for the signatories
y = 0.8

# Set the line properties
line_x_start = 0.1
line_x_end = 0.9
line_y = y - 0.1

# Set the title properties
title = 'Signatories'
title_x = 0.5
title_y = line_y - 0.1

# Add the signatories' names
for name in names:
    ax.text(0.1, y, name, fontsize=12)
    y -= 0.1

# Add the line underneath the signatories' names
ax.plot([line_x_start, line_x_end], [line_y, line_y], color='black')

# Add the title below the line
ax.text(title_x, title_y, title, ha='center', va='center', fontsize=14, fontweight='bold')

# Remove axis and axis labels
ax.axis('off')

# Show the plot
plt.show()
