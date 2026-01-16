# Image Assets for Innovation Hub Plugin

This directory should contain the following image assets:

## Required Images

### 1. background.png
- **Purpose**: Background image for the agent interface
- **Recommended Size**: 1920x1080px or larger
- **Format**: PNG with transparency support
- **Color Theme**: Should complement the #cecece primary color

### 2. logo.png
- **Purpose**: Agent logo displayed in the interface
- **Recommended Size**: 256x256px (square)
- **Format**: PNG with transparency
- **Style**: Clean, professional, innovation-themed

### 3. thumbnail.png
- **Purpose**: Plugin thumbnail shown in the agent selection menu
- **Recommended Size**: 200x200px (square)
- **Format**: PNG
- **Style**: Should be recognizable at small sizes

### 4. Innovation Showcase Images
Additional images to showcase innovations (referenced by the agent):
- Use descriptive filenames (e.g., `ai-analytics.png`, `cloud-platform.png`)
- **Format**: PNG or JPG
- **Recommended Size**: 1200x800px or similar aspect ratio
- **Optimization**: Keep file sizes reasonable for web delivery (<500KB each)

## Image Guidelines

- Use high-quality images that represent innovation and technology
- Ensure images are professional and on-brand
- Consider using consistent visual style across all images
- Optimize images for web (compress without losing quality)
- Test images at different screen sizes

## Example Image References

When the agent responds with image URLs, they should reference images in this directory:

```json
{
  "type": "image",
  "url": "/plugins/innovation-hub/image/your-image-name.png",
  "title": "Image Title",
  "description": "Brief description"
}
```

## Creating Placeholder Images

Until production images are ready, you can create placeholder images with:
- Simple gradients in the #cecece color scheme
- Text overlays indicating the image purpose
- Abstract technology-themed graphics
