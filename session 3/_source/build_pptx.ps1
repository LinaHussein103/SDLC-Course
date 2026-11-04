# Builds 01_Codeflow_UIUX_Session_Slides.pptx from the per-slide PNGs in _source/png using PowerPoint COM.
$src = $PSScriptRoot
$out = Join-Path (Split-Path $src -Parent) "01_Codeflow_UIUX_Session_Slides.pptx"

$notes = @(
  "Welcome. Today continues the SDLC design phase: last session we wrote requirements; today we turn them into interfaces people can actually use.",
  "Ask the room before advancing. Collect 4-6 answers on the board. Don't correct yet - the next three slides will.",
  "UI is the 'what': everything you could point at in a screenshot. Walk through the numbered callouts on the shuttle app.",
  "UX is the whole journey toward a goal. Point at the dip at 'Pay': one bad moment can ruin an otherwise good product.",
  "Same task, two apps. A wins a screenshot contest; B wins with real users. Ask: which would you rather use when late for a lecture?",
  "Ask the room before advancing. Push students beyond 'it looks nicer' - think cost, errors, accessibility, adoption.",
  "The registration portal is functionally correct but unusable. Passing tests is not the same as meeting user needs.",
  "Eight principles, each linked to a Nielsen heuristic or WCAG. Ask students for one example of each from apps they use.",
  "Ask the room before advancing. Let them argue both sides; the answer is 'it depends on what the user is trying to do'.",
  "Contrast system-first vs goal-first flows. Stress the last column: authentication is still required for protected actions.",
  "UX is not a one-time phase. Testing often sends teams back to flows or even requirements.",
  "A user flow is about what the user does - not tables (schema) or servers (architecture). Trace the 'no' branch to the waitlist.",
  "Wireframes show structure, hierarchy, layout and navigation - deliberately without colour or polish.",
  "Low-fidelity = fast and cheap. Encourage students to sketch many ideas before committing to one.",
  "High-fidelity adds real type, colour, components and states. A mockup only becomes a prototype when it simulates interaction.",
  "Compare purpose, detail and interactivity. Quick check: is a beautiful Figma frame with no links a prototype? (No.)",
  "Tools change; the workflow transfers. Explore on paper, map flows in FigJam, design in Figma, test with users.",
  "Mobile-first forces prioritisation. Emphasise that 'most students use phones' is a hypothesis we would validate, not a fact.",
  "Concrete do/don't pairs. WCAG 2.2 target-size minimum is 24x24 CSS px; platform guides recommend about 44-48.",
  "Introduce the case study and the four roles. Hand out 02 (SRS) and 03 Parts A-B only.",
  "Build suspense: the system meets the requirements on paper. Do not reveal any of the problems.",
  "Pairs, 15-20 minutes. Circulate and push pairs to justify each problem with a user need or requirement.",
  "Use these questions to unblock pairs who are stuck. Then run the reveal from handout 03, Part C onward.",
  "Close: design decisions must be justified by user needs, requirements and usability evidence - not taste."
)

$pp = New-Object -ComObject PowerPoint.Application
$pres = $pp.Presentations.Add(0)          # 0 = no window
$pres.PageSetup.SlideWidth = 960          # 13.333 in
$pres.PageSetup.SlideHeight = 540         # 7.5 in
$pngs = Get-ChildItem (Join-Path $src "png\s*.png") | Sort-Object Name
$i = 0
foreach ($png in $pngs) {
  $i++
  $slide = $pres.Slides.Add($i, 12)        # 12 = ppLayoutBlank
  $pic = $slide.Shapes.AddPicture($png.FullName, 0, -1, 0, 0, 960, 540)
  $pic.AlternativeText = "Slide $i"
  $slide.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text = $notes[$i - 1]
}
if (Test-Path $out) { Remove-Item $out -Force }
$pres.SaveAs($out, 24)                     # 24 = ppSaveAsOpenXMLPresentation
$pres.Close()
$pp.Quit()
Get-Item $out | Select-Object Name, Length
