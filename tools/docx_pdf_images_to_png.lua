-- For DOCX output, point PDF figure images at generated PNG previews.
-- LaTeX continues to use the PDF figures; this filter only affects Pandoc's
-- document model while writing Word files.

function Image(img)
  if FORMAT ~= "docx" then
    return img
  end

  if img.src:lower():match("%.pdf$") then
    local basename = img.src:match("[^/\\]+$") or img.src
    img.src = basename:gsub("%.pdf$", ".png")
  end

  return img
end
