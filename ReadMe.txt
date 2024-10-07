--------------
- Readme.txt -
--------------
   This text document was designed to instruct the reader on how to understand and use the Arrow Dynamic Link Library, a product of Digitronix Inc. This file is best read with Word Wrap.

--- -----------------------------
1.0 Introduction to the Arrow DLL
--- -----------------------------
   After the making of Digitronix Tank '99, I saw it was possible to create a rotatable ordered array of points connected together. I began and ended my work on the project to create such just today, which is Tuesday, 13 June, 2000.
   There are a few uses for this DLL that I can think of. First of all, it would in my view be extremely easy to make a guage for a clock or something like what you would find a plane panel in Microsoft Filght Simulator. It is also ideal for games requiring a rotating object.

--- -----------------------------------------
2.0 How to create rotatable graphic functions
--- -----------------------------------------
   If you know how one would plot lines on an x and y axis, you would understand how this control is opperated. Supposed the center of rotation is at the point [0,0] on a graph, and you wanted to create a rotatable graphic from that point. You just pull out a piece of graph paper, plot out the points which will serve as bounderies for your shape, and just draw lines connecting the dots, and list all the ordered pairs that create the shape. Understand, though, that the program reverses the order of the x axis, so that what would otherwise be an arrow pointing up would be pointing down.
   To make the formula for the rotatable graphic, you first derive the ordered pairs, and you list them with the following method.
1) Include all the ordered pairs in a row inside brackets ([]).
2) Determine all the ordered pairs that are to be connected, and connect their brackets to each other with the "-" character.
3) If there is to be no line connecting one ordered pair from the next one, separate their brackets with the ":" character.
   With this method, the user is allowed flexibility in the shapes that they create with this control.

--- -----------
3.0 Programming
--- -----------
To access the DLL, you must first set the ArrowContainer property of the class (i.e. "Set TestPoint = Form1") Otherwise, there would not be enough information to which container to draw the shape. By default, the ArrowString property of the class is set to "[0,4]-[-2,-3]-[0,-2]-[2,-3]-[0,4]", which is an arrow pointing downward. To resize the arrow, you would adjust the ArrowZoom property. The arrow's radius is multiplied by the ArrowZoom property where 1 is exact to the scale mode. To rotate the graphic, you would use the ArrowRotate property, specifying the number of degrees to turn the arrow counterclockwise to it's normal position. Use the ArrowPosX and ArrowPosY properties to set where the arrow is to be drawn, or in a shorter method, the SetArrowPos proceedure. To disable the class, set the ArrowDisabled property to false. This causes the class to refuse to draw the shape.

--- -------
4.0 Contact
--- -------
   Please let me know how you like my submission, and support me with your vote for code of the month. If there are any further questions on your part about the Arrow class, I am available at Digitronix@hotmail.com.
   Visit my webpage at http://www.geocities.com/Digitronix.
   Add me to my buddy list if you enjoy discussing VB issues or have a question you desire to ask me. I am Digitronix82.