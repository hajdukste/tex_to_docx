-- Move citeproc's generated References section before manuscript backmatter.
--
-- Pandoc's LaTeX reader does not use \bibliography{...} as a placement
-- marker for DOCX output. With --citeproc, the bibliography is therefore
-- appended to the end of the document. This filter puts it before the first
-- common backmatter heading so the Word file follows the TeX/PDF order.

local backmatter = {
  ACKNOWLEDGEMENTS = true,
  ACKNOWLEDGMENTS = true,
  ["AUTHOR CONTRIBUTIONS"] = true,
  ["DECLARATION OF INTERESTS"] = true,
  ["SUPPLEMENTARY MATERIALS"] = true,
}

local refs_blocks = nil

local function is_refs_div(block)
  return block.t == "Div" and block.identifier == "refs"
end

local function is_references_header(block)
  return block.t == "Header"
    and pandoc.utils.stringify(block.content):lower() == "references"
end

local function is_backmatter_header(block)
  if block.t ~= "Header" then
    return false
  end

  local text = pandoc.utils.stringify(block.content)
  text = text:gsub("%s+", " "):gsub("^%s+", ""):gsub("%s+$", "")
  return backmatter[text:upper()] == true
end

function Pandoc(doc)
  local without_refs = {}
  local i = 1

  while i <= #doc.blocks do
    local block = doc.blocks[i]

    if is_references_header(block) and doc.blocks[i + 1] and is_refs_div(doc.blocks[i + 1]) then
      refs_blocks = { block, doc.blocks[i + 1] }
      i = i + 2
    elseif is_refs_div(block) then
      refs_blocks = { pandoc.Header(1, "References"), block }
      i = i + 1
    else
      table.insert(without_refs, block)
      i = i + 1
    end
  end

  if refs_blocks == nil then
    return doc
  end

  local moved = {}
  local inserted = false

  for _, block in ipairs(without_refs) do
    if not inserted and is_backmatter_header(block) then
      for _, ref_block in ipairs(refs_blocks) do
        table.insert(moved, ref_block)
      end
      inserted = true
    end
    table.insert(moved, block)
  end

  if not inserted then
    for _, ref_block in ipairs(refs_blocks) do
      table.insert(moved, ref_block)
    end
  end

  return pandoc.Pandoc(moved, doc.meta)
end
