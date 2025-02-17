Office.onReady(() => {
  PowerPoint.createPresentation().then((presentation) => {
    let slide = presentation.slides.add();
    slide.title = "Welcome to OfficeJS";
    slide.content.addText("This is a sample slide deck using OfficeJS.", {
      x: 50,
      y: 100,
      width: 400,
      height: 100
    });
    
    let secondSlide = presentation.slides.add();
    secondSlide.title = "Features";
    secondSlide.content.addText("- Create and edit PowerPoint slides\n- Integrate with Office 365\n- Automate slide generation", {
      x: 50,
      y: 100,
      width: 400,
      height: 200
    });
    
    presentation.save();
  });
});
