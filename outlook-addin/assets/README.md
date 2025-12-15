# Assets Folder

## Required Icons

You need to create the following icon files for your Outlook add-in:

### Icon Sizes Required:

1. **icon-16.png** - 16x16 pixels (Ribbon button small)
2. **icon-32.png** - 32x32 pixels (Ribbon button medium)
3. **icon-64.png** - 64x64 pixels (AppSource listing)
4. **icon-80.png** - 80x80 pixels (Ribbon button large)
5. **icon-128.png** - 128x128 pixels (High-resolution AppSource)

### Icon Design Guidelines:

- **Simple and recognizable** - Should work at small sizes
- **Transparent background** (PNG format)
- **Clear symbol** - Consider using an envelope with an "X" or unsubscribe symbol
- **Brand colors** - Use your app's color scheme
- **High contrast** - Visible on light and dark backgrounds

### Tools for Creating Icons:

1. **Figma** (Free, web-based) - https://figma.com
2. **Canva** (Free tier available) - https://canva.com
3. **Adobe Illustrator** (Professional)
4. **GIMP** (Free, desktop) - https://gimp.org
5. **Photoshop** (Professional)

### Quick Icon Template:

You can use these emojis as temporary placeholders:
- 📧 (Envelope)
- ✉️ (Envelope with arrow)
- 🚫 (Prohibited)
- ✖️ (X mark)
- 📭 (Closed mailbox)

### Example Icon Concept:

```
Simple design idea:
- Blue envelope icon
- Red "X" or "minus" symbol overlaid
- Clean, modern look
- Works at 16px
```

### Generating Icons from SVG:

If you have an SVG icon:

```bash
# Install ImageMagick
sudo apt-get install imagemagick

# Convert SVG to multiple PNG sizes
convert icon.svg -resize 16x16 icon-16.png
convert icon.svg -resize 32x32 icon-32.png
convert icon.svg -resize 64x64 icon-64.png
convert icon.svg -resize 80x80 icon-80.png
convert icon.svg -resize 128x128 icon-128.png
```

### For Now (Testing):

You can use placeholder images from:
- https://via.placeholder.com/16x16
- https://via.placeholder.com/32x32
- https://via.placeholder.com/64x64
- https://via.placeholder.com/80x80
- https://via.placeholder.com/128x128

Or use a simple colored square as a temporary icon during development.

### Microsoft Design Guidelines:

For official icon guidelines, see:
https://docs.microsoft.com/office/dev/add-ins/design/add-in-icons
