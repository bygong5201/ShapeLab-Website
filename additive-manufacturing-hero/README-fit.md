# Additive Manufacturing Parallax Background Fit Pack

This pack is meant to be dropped into an existing website that already has an `index.html`.
It does not replace your main file.

## Files

- `shape-lab-background-fit.css` - CSS override that sizes the blueprint background to the full website viewport.
- `shape-lab-background-fit.js` - optional pointer/scroll parallax motion.
- `assets/additive-blueprint-bg-1920.png` - optimized 16:9 background for normal screens.
- `assets/additive-blueprint-bg-2560.png` - high-resolution background for wide screens.

## Add to your existing `index.html`

Place this inside `<head>` after your main stylesheet:

```html
<link rel="stylesheet" href="additive-manufacturing-hero/shape-lab-background-fit.css">
```

Place this immediately after the opening `<body>` tag if your page does not already include `.parallax-scene`:

```html
<div class="parallax-scene" aria-hidden="true">
  <div class="parallax-layer" data-parallax-speed="1"></div>
</div>
```

Place this before `</body>`:

```html
<script src="additive-manufacturing-hero/shape-lab-background-fit.js"></script>
```

## Notes

The CSS uses a fixed full-viewport scene, `100dvh`, responsive background sizing, and mobile-specific positioning so the background fills the website without awkward cropping.
