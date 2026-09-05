function Pandoc(doc)
  local firstMainSection

  for index, block in ipairs(doc.blocks) do
    if block.t == "Header" and block.level == 2 then
      firstMainSection = index
      break
    end
  end

  if not firstMainSection then
    return doc
  end

  local blocks = {
    pandoc.RawBlock("latex", "\\begin{resumeheader}")
  }

  for index = 1, firstMainSection - 1 do
    table.insert(blocks, doc.blocks[index])
  end

  table.insert(blocks, pandoc.RawBlock("latex", "\\end{resumeheader}"))

  for index = firstMainSection, #doc.blocks do
    table.insert(blocks, doc.blocks[index])
  end

  doc.blocks = blocks
  return doc
end
