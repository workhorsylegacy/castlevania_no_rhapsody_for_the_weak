VERSION 5.00
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H0080FF80&
   BorderStyle     =   0  'None
   Caption         =   "No Rhapsody for The Weak"
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\
' Castlevania : No Rhapsody for The Weak
' Main Form
' Programmed by Matt Jones
' Started on : 10/09/02
' Last update : 07/30/03
' -------------------------------------------------------------------
' GENERAL DECLARATIONS
'\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/


'--------PROGRAM NOTES--------'
'Put all loops inside functions so they don't have to be called for each check of the loop.
'Make xxxCol.txt files for the projectiles to use on weapons Alucard has hit back at foes
'Give projectile weapons 50 extra pixels to move off screen before they are destroyed
'Sinking trough platforms does not work with new state machine!
'Add spark to candle's destruction
'Add items such as hearts to itmSpr.gif so they can be grabbed from candles. Have the first
'few items pickups things like hearts & money that arent added to the inventory
'
'
'Get a good Paint program and fix three layer backgrounds!!!!!!!!!!
'Have movement code for diagnols!!!!!!!!!!!!
'Use getascii key thingy instead of keydown because it is not working with a ton of keys down.
'!!!!!!!!!!!!!!!Update sword collision file because it is outdated!!!!!!!!!!
'Add Warg character.

'Remove cape and hair from collision
'Resize collision arrays to be smaller
'Shrink State machine with < & > signs
'Replace background scrolling
'Add pushing animation
'Add remaining animations
'07/30/03-Most of the OOP translation is complete. Only minor objects
'         such as projectiles and water need to be converted. More
'         with statements need to be used along with the background
'         changing code, and most references to (CH) need to be gone.
'07/28/03-Shrank the size of the collision arrays by using the
'         ammount of lines in the text col files. Started conversion
'         to OOP for much easier use of variables.
'07/24/03-Gave memScreen the value of me.hdc, and replaced me.hdc
'         with memScreen to not have to refer to the object. Fixed
'         error with attack animations being messed up when jump is
'         pushed. Fixed error with EndRez, that stopped it from
'         trying to change the resolution when the resolution was not
'         changed in the begining.
'07/09/03-Greatly improved speed by having the functions that do
'         movement for characters contain the loops for each
'         character rather than having the function called each time
'         by each character. This means that the function is only
'         called once each loop, rather than once for each character
'         in the loop. Started basic efects for the candles when
'         destroyed. Simplified the way the game loop checked for a
'         game over and the end of the program.
'07/08/03-Fixed the main loop so it would flip the buffer to the
'         screen on frames that haven't been updated, rather than
'         redrawing the entire buffer piece-by-piece. Fixed
'         resolution changing code to actually work - Maybe.
'07/04/03-Fixed problem with items being replaced by incorrect items
'         when the list of exposed items was being updated. An
'         inventory array was incorrecly nested with an item array.
'07/02/03-Added real water splashing effects. It still only works
'         once though. Fixed problem with items not having the right
'         weight for their gravity.
'06/20/03-Fixed zombie so it hibernates when too far off the screen
'         and re-spawns when Alucard is in range. Made projectile
'         weapon collision detection work. Fixed error that would
'         only let zombies throw projectiles when the background had
'         an X of zero. Fixed error with projectile weapons being
'         drawn incorrectly when facing left. Fixed error with the
'         collision detection for the projectiles getting stuck on
'         by not resetting the collision value. Fixed error in the
'         projectile collision detection that had some left over
'         arrays from the item collision detection it was copied
'         from. Fixed problem with the code that delets and replaces
'         projectiles weapons that updated everything but the
'         direction the projectiles was traveling. This would result
'         in projectiles sudenly changing direction in mid-air.
'06/17/03-Fixed problem with sub weapon animations not being replaced
'         properly when an axe is deleted off the screen.
'06/16/03-Made sub weapon axes work. They can be thrown until they
'         move off the screen where they are deleted. No collision
'         detection, gravity, or weapon bounding(hit projectiles with
'         sword, and maybe back at enemy).
'06/14/03-Made gravity for jumping. Distance moved decreases every
'         time Alucard moves up. The distance is reset to normal when
'         Alucard's gravity detects a collision with the ground.
'06/13/03-Fixed error where gravity would move Alucard, then when he
'         started falling and using his sword for the first frame
'         he would move again. Re-arranged code so collision and
'         coordinates files for Items, headsup display, and Alucard
'         would be loaded with the program rather than from the save
'         file, because they never change.
'06/11/03-Fixed problem with the state machine letting falling
'         attacks be processed when Alucard was landing. This would
'         result in him getting stuck in his sword swinging
'         animation. Also fixed the problem with background layer
'         three and four being drawn when resetting to a save room or
'         room with only one background layer.
'05/29/03-Replaced Alucards state machine. It now works mostly off of
'         what button was pushed, and is faster. There are a few
'         errors with jumping and sword swinging while standing.
'05/06/03-Fixed small error that had the LoadDC code deleting the
'         temporary objects before it had a chance to get the image
'         dimentions. This would leave the objects holding the
'         dimentions in memory, causing memory leaks. Made it so
'         there could be a third background layer. All that's needed
'         is for the XX-X-EX.TXT files to list the images, and the
'         program will take care of the rest.
'05/05/03-The sword collision file was updated to work with falling
'         while slashing down and out. Fixed small error that would
'         cause Alucard to only jump a few pixels high after using
'         his sword in the air. It was caused by the fall sword out,
'         fall sword down, and jump sword animations setting the jump
'         distance to its max when they finished (In hopes that
'         Alucard would not be able to jump again). Got rid of the
'         separate sword animation routines, and replaced them with a
'         function. Fixed error that allowed Alucard's state to be
'         set to stop, while in the air. This allowed you to
'         repeatedly push the direction and attack, to repeatedly
'         jump and slash up in the air. Made ChangeEnemies reset the
'         states and frames of animations to their starts as each is
'         loaded. Fixed an error that caused the enemy movements to
'         use their distance and not their space to calculate
'         collisions. Made zombie character more like SOTN. It will
'         spawn from a designated point, unspawn when out of range,
'         and turn around if it hits walls or stairs.
'04/30/03-Added animation with Alucard having his legs held together
'         while standing still. Chopped last animation off on
'         standing up animation. Made teetering animation work when
'         stopping instead of standing. Made still animation work
'         when standing up, or landing, or stopping on surfaces that
'         cause teetering. Made last three frames of the falling
'         sword down and out animations stick for three frames,
'         because they are too short. Added falling attacks to call
'         for ActionAirHurt to eliminate breaks in hurt machine.
'04/29/03-Made ActionTurnAround set the new direction before it moved
'         Alucard to fix the glitch that moved the opposite direction
'         in the first frame of animation. Fixed having the wrong
'         frame of animation at the begining of turning when comming
'         from walking, by having the ActionTurnAround function set
'         the first frame of animation to the first frame of turning
'         rather than one before (because of the delay it would not
'         be updated). Made turn restart when turning back-to-back in
'         different directions.
'04/27/03-Jumping and falling sword animation and collision detection
'         is all but perfect. There is a small error with it not
'         always working when not holding the direction, and another
'         when you try to jump after slashing and falling.
'04/26/03-Nearly finished the jumping and falling sword animations.
'         They just need some work to align the swords to Alucard and
'         make the state machines more stable.
'04/23/03-Made landing cancel out sword air out to fix sticking
'         state.
'04/20/03-Fixed errors with items being in the wrong place after
'         other items have been removed or added. Added jumping sword
'         animation. Needs falling sword. More work should be done on
'         having landing and hard landing cancel out animations.
'04/16/03-Items can now be recieved by touch and added to inventory.
'         There is a glitch that swaps the placement of items when
'         items are suppose to be deleted.
'04/14/03-Fixed the sprite section used for Alucard's jumping
'         animation. His hair now moves more realistically. Made it
'         so Alucard's collisions with items could be detected. All
'         that needs to be done is for them to be erased from the
'         screen and added to the inventory array.
'04/09/03-Fixed sloppy transition between turning and walking
'         animation by starting the walking animation half way
'         through, where the cape, hair and legs are matching.
'04/09/03-Tweaked Alucard's animations: Turning around now takes
'         twice as long. Falling down through the air is now in
'         alignment with the rest of sprites. His feet now touch
'         the ground while walking (even though they don't in
'         SOTN -- his head controlled walking alignment). Fixed
'         small error with turning around having Alucard face the new
'         direction before he started turning. Fixed similar error
'         with jumping animation not updating the state until after
'         the first delay. Replaced the form's surface as the screen
'         and eliminated all remaining picture boxes. Replaced some
'         VB color constants with real values for speed.
'04/07/03-Fixed problem with item collision detection. It now works
'         with the correct dimensions, gravity, collision detection,
'         items, and source objects. Added weight to items so they
'         can fall at different speeds.
'04/06/03-Item gravity and collision detection complete. For now all
'         items fall straight down. There is an error with the
'         collision detection knowing the dimensions of the sprites,
'         which causes collision errors.
'04/05/03-Made candles that are destroyed turn into items. Basic work
'         is done. They need collision detection, falling gravity,
'         and an inventory/menu to be added to and viewed from.
'04/01/03-Made it so candles and other destroyable objects could be
'         wiped away with a hit of the sword. They need proper
'         destruction animations and some item birth giving code when
'         destroyed. The ability to use the sword in the air would
'         also help greatly!
'03/25/03-Sped up the loading of images by only getting the width and
'         height of the background images instead of all of them.
'03/24/03-Removed the need to use pictureboxes when loading which
'         speed things up for loading and screen drawing. Also
'         tweaked the sound to keep the same song, load a new one,
'         or not play anything because there is no file, depending
'         on what is saved in the MU.txt file.
'03/22/03-Fixed problems with large backgrounds causing slowdown.
'         This was accomplished by eliminating the pictureboxes and
'         loading real images directlt into the memory DCs rather
'         than creating them as a virgin bitmap and then passing the
'         image in.
'03/15/03-Added more rooms to the castle entrance and tweaked how
'         some of them load new backgrounds.
'03/14/03-Made it so the program can eat text files that tell what
'         sound effects to play for Alucard, Enemies, and the
'         background music. Support for .MP3 files is needed because
'         .WAV take up too much space.
'03/13/03-Added support for sound effects and music with DirectX 8.
'         The code will be moved to the end of the background
'         changing code and open an ENT-1-MU.txt file that points to
'         the music and sound effect files to load. Fixed stupid error
'         with the character's sprite being out of alignment when
'         facing left. The problem was caused because I was using
'         Alucard's coordinates to calculate all the characters left
'         Xpos.
'03/11/03-Fixed the collision boxes for the enemy characters. They
'         now will collide with the backgrounds properly. Walking
'         left for all the characters is working, except for a
'         strange glitch: The Xpos for mirrored sprites (facing left)
'         does not have any effect on its sprite. All tests prove
'         otherwise, but the error still exists.
'03/05/03-Added the BloodyZombie character and tweaked the enemy code
'         to make it easier to have multiple enemies share the same
'         sprites and coordinates.
'03/02/03-Fixed error where the program would continue to keep the
'         same number of background layers when Restarting after a
'         Game Over screen. This would create a phantom image of the
'         old background, because there were blank layers being
'         drawn. Scrolling of background layers is now perfect.
'02/26/03-Added beginning of the heads-up display. The health numbers
'         and health bar work. Vertical scrolling of the second
'         background layer is wrong.
'02/25/03-Added more backgrounds and redid the naming conventions for
'         the files and images to have a common scheme.
'02/21/03-Fixed state machine to work correctly while ducking and
'         attacking. This includes both downward and outward swings.
'         Replaced many instances of ANI(CH) with W to speed up
'         animation routines and screen drawing. Similar things were
'         done through out large occurrences of nested variables.
'02/20/03-Added slashing forward while ducking animations. The
'         slashing down does not work yet, but it's done.
'02/19/03-Fixed game over screen so you could start a game over from
'         scratch by dumping and reloading all the variables and
'         memory DCs. Made it so if you push ESC then you go straight
'         to the Game Over screen. Made Alucard's outline
'         transparent. Made the SCollisionDetection so it doesn't
'         check for weapon collisions unless DrwWep(CH)=true. This
'         eliminates some collision detection errors at close range.
'02/17/03-Added a Game Over screen. Restarting a game needs some work
'         though. Tweaked moveAlucard by making it into a function and
'         shrinking down some variables.
'02/15/03-Fixed minor error with damage showing on the outside of
'         Alucard as a pink box, instead of over him.
'02/13/03-Worked feverishly on the memory DC problem with older
'         versions of Windows(95, 98, ME). Might have fixed problem.
'         Program now runs a little faster thanks to memory DCs only
'         being created before the game starts and not every time
'         a collision takes place. Some of the state machines were
'         sped up by using small named variables in place of nested
'         arrays.
'02/12/03-Changed background layer scrolling to be based on the
'         position of the first background layer. As in
'         BGX2 = BGX * 2 \ 3. This will insure proper placement and
'         eliminate the chance of it scrolling too far in one
'         direction. This also means that Alucard can move any amount
'         of pixels he wants in one loop without having to move the
'         same in the other direction to fix the background. Tweaked
'         jumping animation to start when the delay starts. Added the
'         other jumping damage animations.
'02/11/03-Removed a bunch of slow constants. Started to speed up the
'         state machines by using select case. Started to speed up
'         memory DC drawing by using 0 as a source and vbWhiteness
'         as a dwRop. The collision images need to have their colors
'         inverted for it to work though.
'02/08/03-Used getDC API call instead of frmMain.hdc to try and fix
'         memory leak problem. Sped up Dir checking with Dir$.
'02/07/03-Probably finished tweaking pixel collision detection. Made
'         each background have its own backgroundCh.txt file, that
'         will determine character starting position, enemy count
'         and which enemies are used.
'02/05/03-Started major optimizations. Used Len() instead of ="" for
'         strings. Used \ instead of / for division. Replaced most
'         local integer variables with long variables. Added support
'         for multiple enemies. Tweaking pixel collision detection.
'02/04/03-Removed gravity, but you still hit the ground hard when
'         falling from great heights, or when you get hit by an
'         enemy.
'02/03/03-Fixed minor problem with enemies being out of place in the
'         pixel collision detection when facing left. Fixed background
'         scrolling to move the layers the same distance, as to not
'         slowly move them more and more off the screen. Added zombie
'         death animation. Maybe fixed the memory DC problem by having
'         it select the previous object before deleting the DC.
'01/31/03-Fixed the pixel perfect collision detection. Now you can kill
'         the zombie.
'01/30/03-Eliminated the need to use picture boxes for image storing.
'         Memory DCs are standard now. This slightly speeds up screen
'         drawing. Made Alucard bounce off a wall if an enemy hits him
'         into it. This completly eliminates him getting pinned
'         between a foe and a wall.
'01/29/03-Added support for background layers. They also scroll along
'         the correct axis and correct distances.
'01/28/03-Fixed resolution changing problem by not using the
'         SendMessage API call that it originally used. This
'         would only have been necessary if we were also going to
'         change the size of windows. Also fixed the problem that
'         caused the pieces for the background collision's unpassable
'         objects to be drawn in the wrong place if you enter the
'         screen from the left. Fixed animation sticking problem when
'         you jump straight up and then hold the opposite direction
'         you are facing. Also fixed the sticking frames when turning
'         around too fast, or holding both right and left.
'01/26/03-Fixed the background changing to work perfectly for any
'         direction. More backgrounds and exit files can simply be
'         plugged right into it. The memory DC seems to cause
'         problems in Windows ME.
'01/25/03-Created two more frames of animation for the right and left
'         jumping. This boost the number from 4 to 6 and makes it
'         less repetitive when it repeats the last 4 frames in the
'         end (was only 2). Fixed last frame of landing animation.
'         Sped up environmental collision detection by using a memory
'         DC as a scratch pad. Figured out a better way of getting the
'         Left Xpos (for mirroring sprites) using the character's
'         center (CC) which is always going to be the same.
'         Left Xpos = CC - (Width + RightXpos - CC).
'01/23/03-Made Alucard move in a random direction when hit by an
'         enemy. This will stop him from getting pinned between foes
'         and walls. Also went over the code and tried to minimize
'         everything and speed it up. Updated the background loader
'         to not keep the previous background's objects when there
'         are no new ones - moving from a background with candles and
'         rocks to a background with nothing, like a save room.
'01/22/03-Added the basic hurt animations for Alucard and hooked them
'         to the collision detection so they work. The collisions in
'         the air only have one animation right now. Added basic
'         pink coloration to Alucard when he is hurt.
'01/18/03-Fixed the soft platforms so when you jump up through them
'         and land, it doesn't pull you up when you don't jump high
'         enough. Fixed jumping animation's frames. Fixed ducking and
'         standing animation frames.
'01/17/03-The dithering color problem was fixed. It was caused when
'         the bitblt call dithered the colors in 16-Bit color mode.
'         Both collision detection schemes were sped up immensely by
'         by using Select Case instead of If statements. The
'         background collision images were also simplified by no
'         longer using orange for ledges and just letting it
'         calculate that.
'01/16/03-Completed more work on background changing. Made it switch
'         the backgrounds, as well as the passable and unpassable
'         objects and their collisions. The permanent slowdown was
'         also eliminated while changing backgrounds. The remaining
'         jumping animations were also added to the straight up jump.
'01/15/03-Fixed the background scrolling and implemented a basic save
'         file. Most of the other stuff involving the background
'         loading was redone. The BackgroundScroll now can tell the
'         difference between vertical, horizontal, or free scrolling
'         backgrounds and works accordingly.
'01/13/03-Pixel collision detection is now in perfect working order.
'         All it needs now is some damage animation and sprite sheets.
'         Alucard Pixel collision sprite sheet finished. Developed a
'         method of making them in photoshop by layering two of the
'         same masks on top of each other and moving one down a
'         pixel, up a pixel, right a pixel, and left a pixel. The
'         exposed uderlayer is saved and the hanging-over top layer
'         is removed to create a perfect outline. Added teetering
'         animation to fix foot hang on ledges and staircases. Also
'         added the twisting hurt animation for huge amounts of
'         damage and death.
'01/10/03-Completed more work on the pixel perfect collision. Made it
'         scan from left to right then move down to be faster than
'         scanning up to down then moving right.
'01/09/03-Made Soft and mblnSink into Soft(CH) and mblnSink(CH) so the
'         characters movements wouldn’t effect the softplatform
'         collisions of other characters. Added four more frames
'         of animation to the candles to bring it up to 6.
'01/07/03-Moved the generation of the left Xpos and weapon collision
'         detection into its own sub that is called after everything,
'         but just before the screen draw. Fixed pixel perfect
'         collision detection. The loop was scanning a pixel too far
'         to the right and may be off by one on the left, top, or
'         bottom too.
'01/04/03-Used IIF functions instead of multiple IFTHEN statements
'         when inside an animation sub, to reduce the size immensely.
'         Started memory DC work for the pixel perfect collision
'         detection.
'01/03/03-Moved the candle state machine into the effect section. It
'         is also called from the moveEffect sub with all other
'         effects.
'12/30/02-Added support for platforms that you could jump through
'         the bottom and land on the top, or jump while ducking and
'         fall through completely.
'12/29/02-Made standing animation move on to bobbing. Made hard
'         landing animation use the ending of the ducking sword
'         animation. Changed other landing, jumping and standing
'         animations to not skip their first frame of animation.
'         Added a delay to bobbing animation to match real speed.
'         Also added the same delay to zombie animation to slow it
'         down.
'12/27/02-Added a hard landing code for when falling from great
'         heights. It currently uses the ducking animation and then
'         has Alucard stand up automatically. Also tweaked the
'         collision detection to work perfectly, as long as the floor
'         is at least 5 pixels thick.
'12/26/02-Simplified the code for moving right, so it would be as
'         short as the code for moving left. Made all the characters
'         use the same variables for their position and the same code
'         for collision detection.
'12/23/02-Setup program for having all the enemy and character
'         collision and variable code the same.
'12/20/02-Fixed right and left collision detection to work correctly
'         and eliminated the bouncing problem when hitting varying
'         height and shaped walls. Also fixed the jumping animation
'         so gravity would take over when a ceiling was hit. The mask
'         was also updated because the tolerance on the old one was
'         off.
'12/19/02-Painstakingly redid the sword effects. Created a new method
'         that will be used on all the other sword effects as a
'         template, making it much easier. Also fixed the resolution
'         changer and fixed an error that allowed the sword animation
'         to be canceled out when any non game key was pressed, like
'         a, ctrl, or f3. Also fixed the bouncing in and out of the
'         sides of small protruding objects. Added a small picturebox
'         to the opening screen that detects crappy video cards.
'12/18/02-Re-did most of the collision detection to be very fast. It
'         now only needs to be configured to move subtle distances
'         that happen when a movement is interrupted by a collision.
'12/16/02-Fixed the problem with characters sinking into the floor
'         when they move up or down a flight of stairs. It works
'         because it moves them regardless if there is a collision -
'         because they are still standing on stairs allowing them
'         to move up or down.
'12/15/02-Re-coded the collision detection to actually work from
'         where the sprite is located. Also made it accurate to the
'         pixel where it hits the unpassable object.
'12/12/02-Completed most work on the backgrounds' passable and
'         unpassable objects, their masks and collision detection,
'         candle movement, and loading and masks.
'12/11/02-Completed more work on enemy physics, sprites and AI. The
'         sprite placement and some AI routines are all that need to
'         be added.
'12/10/02-Fixed the sword engine so it would self terminate and
'         cancel out the WEPs and drawing if the attack was
'         canceled by movement. Also completed standing sword
'         animation for the right and left sides. Also added some
'         enemy animation. It needs the rest of Alucard's code to be
'         converted and some AI.
'12/09/02-Some of the states were sticking because the sprites
'         weren't changing and the states never got an opportunity
'         to change. This was fixed by changing the state first.
'12/08/02-Completed mirroring of sprites. Also started to fine tune
'         positions and movement.
'12/04/02-Redid sprite sheets and started fixing the animations to
'         use them. Also, came up with a plan to use a mirroring
'         BitBlt for the left side.
'12/02/02-Started sword animations. Also added bobbing animations and
'         made is go off when Alucard is standing still or stops
'         dashing instead of slamming on the breaks.
'11/25/02-Made a direction cancel-out if the opposite is pressed
'         while it is still being held down. This only applies to the
'         right and left for keyDown and keyUp. Also started the
'         drawing of the small objects into the background such as
'         candles and rocks.
'11/23/02-Fixxed the sinking into the floor by having the collision
'         detection note how far stuff sinks and then add that to the
'         BitBlt so it will always be right where it should. It needs
'         to be used in the stairs to stop sinking all together.
'11/21/02-Sped up the collision detection by only having it work in
'         the areas that are red (where Alucard is standing), instead
'         of having it draw and search the entire mask.
'11/20/02-Slightly tweaked the background changing and jumping. Also
'         rearranged the code to be more organized.
'11/19/02-Jumping now works. It seems a little too mechanical and
'         moves in too straight a line, but is works. Also added
'         ghetto support for background changing.
'11/15/02-Fixed the way the screen scrolled while moving every which
'         way on the stairs. Also used a separate color for each
'         type of stair case, right and left. Also started falling
'         movement animations.
'11/14/02-Added support for moving up stairs that actually works.
'         Walking back down them is strange though.
'11/11/02-New collision detection implemented using rectangles
'         instead of pixels. It will be used for environmental
'         collision. The pixel perfect collision detection will be
'         reserved for Alucard VS  Enemy collisions.
'10/25/02-Started jumping animation. There seems to be some
'         confusion on how to pull it off. Especially the collision
'         detection for the ceiling. Also tweaked the turning
'         animation to use only two variables for the SetStart.
'10/24/02-Added animations and constants for standing up for the
'         right and left sides.
'10/22/02-Added a falling and landing animation for the left and
'         right sides. Also added landing states and the current
'         state is now visible in the caption of the collision form.
'10/19/02-Slightly modified the drawing routines and did some speed
'         tests. The CreateDIBitmat and GetDIBitmatTemplate code
'         will speed things up immensely. Hurry up and implement them.
'         Collision detection was also tweaked. It works a lot better
'         with smaller masks and smooth black areas on both of the
'         masks. Also smoothed out Alucard's masks by having his hair
'         and cape a dark gray instead up black.
'10/18/02-Implemented vertical scrolling, both up and down. Gravity
'         now works. As does the collision detection of the floor.
'         There are some problems with the size and shapes of the
'         background masks that cause hang-ups though.
'10/15/02-Added ducking animations for the right and left. Added all
'         the resolution changing code, but it isn't in use because
'         it is not working correctly. A better background BitBlt is
'         now being used. It May cause problems though.
'10/14/02-Added the stopping animation. Updated the SetStart code so
'         if the numbers aren't any of the ones used for the
'         animations they are dumped and replaced by zero.
'10/11/02-Collision detection was speed up considerably by only
'         having it run the CollisionDetect function when the screen
'         or Alucard scrolls. Also had to move masks to their forward
'         three pixels so the cape wouldn’t catch on the walls.
'10/10/02-Picture boxes and basic code for the collision detection
'         are now complete. More basic movement added--dashing right
'         and left, walking left and right and turning around from
'         both sides is now complete.
'10/09/02-Started basic screen drawing routines and sheet loaders.
'         Transferred most of the salvageable code and fixed it up to
'         work more efficiently over here. Some basic movement has
'         been completed without the use of those damn timers.

'--------FUTURE WORK NOTES--------'
'Make the background mask smaller so it cant scroll up any more. Also
'         still have it big enough so the floor is seven pixels thick.
'Use a mask of Alucard that has a gray cape to give his cape a
'         transparent effect rather than the psx's flashing (only
'         when he has The Green Semi-Transparent Cape).
'Try to get rid of any state setting inside anything that isn't an
'         ActionMovementDirection procedure, or it will be a
'         nightmare later on.

'-------GLITCH LOG--------'
'----BROKEN----'
'1.If you jump through a soft platform and land on another soft
'        platform, you will fall right through it. This is because
'        the collision detection is turned off when falling through a
'        soft platform and then reactivated when landing on a hard
'        one. This also means that if you pass through a soft platform
'        while walking, it will deactivate the collision detection
'        causing you to fall until you are out of its range.
'----FIXED----'
'1.When you get close to a wall, you can repeatedly push the same
'        button to hump your way through it one pixel at a time.
'        FIX:The problem was that the ANI was not being reset
'            until after the first collision detection. It was
'            scanning a 0 x 0 size grid and not doing anything until
'            the SetStart function was run, just after the first
'            collision detection.
'2.When you try to move while falling, the state sticks in the last
'       direction you were facing when you hit the last frame of
'       falling animation.
'       FIX:This was fixed by having the state set before ANI
'           was updated, so it would no longer depend on the value of
'           ANI - to get inside the if statement -, only if the
'           animation was called.
'3.If you turn in the opposite direction really fast, or push the
'        opposite direction down while still holding the original
'        button, the animations will stick.
'        Fix:The turning animation included the walking animation's
'        frame numbers in the SetState. This caused it to not update
'        the frame, but update everything else.

Option Explicit

'Dim lngSound As Long
Dim DrawMenu As Boolean
Dim GameOver As Boolean

Dim mblnItems As Boolean ' Items print to screen
Dim ItemLoop As Long ' Used in itrem drawing loop
'Dim FR As Long ' Frame Rate
'Dim F1 As Long, F2 As Long ' Frame Timmer


'/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\
' GENERAL DECLARATIONS
' -------------------------------------------------------------------
' FORM LOAD
'\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/


Private Sub Form_Load()

    Me.Show
    SavePicture frmChangeRez.picScreenBox.Image, App.Path & "\Backgrounds\blank.bmp"
    
    memScreen = Me.hdc
    memBlank = LoadDC("", False)
    memBuffer = LoadDC("", False) ' Create Buffer
    memCollEnv = LoadDC("", False) ' Create Enviornmental Collision Buffer
    memCollCh = LoadDC("", False) ' Create Character Collision Buffer
    'Set DX = New DirectX8 ' Make DirectX 8 Available
    'Set DS = DX.DirectSoundCreate("") ' Make Sound Available
    'DS.SetCooperativeLevel hwnd, DSSCL_NORMAL ' Standard Sound
    CH = 0 ' Alucard
    Player.STA = STANDRIGHT
    Me.Show
    LoadRequiredSprites
    LoadUpSprites
    UpdateBackgroundCollision
    LoadNextBackground
    Me.SetFocus
End Sub


'/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\
' FORM LAOD
' -------------------------------------------------------------------
' MAIN ENGINE
'\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/


Sub GameLoop()
Dim T1 As Long, T2 As Long
Dim AC As Boolean ' Animations complete
    Do Until GameOver = True
        T1 = GetTickCount()
        'If FR = 0 Then F1 = GetTickCount()
        Do Until T2 - T1 >= DELAY
            If AC = False Then ' Redraw The Buffer
                'All Animations
                MoveEnemy
                CH = 0 ' Alucard
                MoveAlucard Player.STA, Player.ANI
                CheckBackground
                FinalCalculations
                AC = True
                DrawScreen
            Else ' Flip The Buffer
                BitBlt memScreen, 0, 0, SCREENWIDTH, SCREENHEIGHT, memBuffer, 0, 0, &HCC0020 ' SRCCOPY
            End If
            T2 = GetTickCount()
            'If FR < 3000 Then FR = FR + 1
            DoEvents
        Loop
        'F2 = GetTickCount()
        'If F2 - F1 > 1000 Then FR = 0
        AC = False
    Loop
    If GameOver = True Then
        GameOverScreen
    End If
End Sub

Private Sub DrawScreen()
Dim Obj As Long, W As Long, P As Long, H As String, DW As Long, T As Long
'Dim R1 As Long, R2 As Long ' Total & Free Memory
'Dim T1 As Long, T2 As Long
    'T1 = GetTickCount()
    'Clear buffer
    If BgLayer(2) = False Or BgLayer(3) = False Then BitBlt memBuffer, 0, 0, 320, 240, memBuffer, 0, 0, &H440328 ' SRCERASE
    'Draw Layer 3
    If BgLayer(3) = True Then
        BitBlt memBuffer, 0, 0, SCREENWIDTH, SCREENHEIGHT, memBgSpr3, -BgX3, -BgY3, &HCC0020 ' SRCCOPY
        BitBlt memBuffer, 0, 0, SCREENWIDTH, SCREENHEIGHT, memBgMsk3, -BgX2, -BgY2, &H8800C6 ' SRCAND
    End If
    'Draw Layer 2
    If BgLayer(2) = True Then
        BitBlt memBuffer, 0, 0, SCREENWIDTH, SCREENHEIGHT, memBgSpr2, -BgX2, -BgY2, &HEE0086 ' SRCPAINT
        BitBlt memBuffer, 0, 0, SCREENWIDTH, SCREENHEIGHT, memBgMsk2, -BgX, -BgY, &H8800C6 ' SRCAND
    End If
    'Draw Layer 1
    BitBlt memBuffer, 0, 0, SCREENWIDTH, SCREENHEIGHT, memBgSpr, -BgX, -BgY, &HEE0086 ' SRCPAINT
    'Draw passable background objects
    For Obj = 1 To TotlCand
        BitBlt memBuffer, BgX + PsXpos(Obj), BgY + PsYpos(Obj), PsWp(Obj), PsHp(Obj), memBgPsMsk, PsXp(mintCandle), PsYp(mintCandle), &H8800C6 ' SRCAND
        BitBlt memBuffer, BgX + PsXpos(Obj), BgY + PsYpos(Obj), PsWp(Obj), PsHp(Obj), memBgPsSpr, PsXp(mintCandle), PsYp(mintCandle), &HEE0086 ' SRCPAINT
    Next
    'Draw Enemies
    For CH = 1 To TotlEn
        With Enemy(CH)
            T = .EType
            W = .ANI
            If .pLeft = True Then
                StretchBlt memBuffer, .PX + .LXP + EWidthp(W, T) + BgX, .PY + EYpos(W, T) + BgY, -EWidthp(W, T), EHeightp(W, T), memEMsk(T), EXp(W, T), EYp(W, T), EWidthp(W, T), EHeightp(W, T), &H8800C6 ' SRCAND
                StretchBlt memBuffer, .PX + .LXP + EWidthp(W, T) + BgX, .PY + EYpos(W, T) + BgY, -EWidthp(W, T), EHeightp(W, T), memESpr(T), EXp(W, T), EYp(W, T), EWidthp(W, T), EHeightp(W, T), &HEE0086 ' SRCPAINT
            Else
                StretchBlt memBuffer, .PX + EXpos(W, T) + BgX, .PY + EYpos(W, T) + BgY, EWidthp(W, T), EHeightp(W, T), memEMsk(T), EXp(W, T), EYp(W, T), EWidthp(W, T), EHeightp(W, T), &H8800C6 ' SRCAND
                StretchBlt memBuffer, .PX + EXpos(W, T) + BgX, .PY + EYpos(W, T) + BgY, EWidthp(W, T), EHeightp(W, T), memESpr(T), EXp(W, T), EYp(W, T), EWidthp(W, T), EHeightp(W, T), &HEE0086 ' SRCPAINT
            End If
        End With
    Next
    'Draw Alucard
    With Player
        CH = 0 'Alucard
        W = .ANI
        If .pLeft = True Then
            StretchBlt memBuffer, .PX + .LXP + .Widthp(W), .PY + .Ypos(W), -.Widthp(W), .Heightp(W), .memMsk, .Xp(W), .Yp(W), .Widthp(W), .Heightp(W), &H8800C6 ' SRCAND
            If .HURT = True Then StretchBlt memBuffer, .PX + .LXP + .Widthp(W), .PY + .Ypos(W), -.Widthp(W), .Heightp(W), .memHurt, .Xp(W), .Yp(W), .Widthp(W), .Heightp(W), &HEE0086 ' SRCPAINT
            StretchBlt memBuffer, .PX + .LXP + .Widthp(W), .PY + .Ypos(W), -.Widthp(W), .Heightp(W), .memSpr, .Xp(W), .Yp(W), .Widthp(W), .Heightp(W), &HEE0086 ' SRCPAINT
        Else
            StretchBlt memBuffer, .PX + .Xpos(W), .PY + .Ypos(W), .Widthp(W), .Heightp(W), .memMsk, .Xp(W), .Yp(W), .Widthp(W), .Heightp(W), &H8800C6 ' SRCAND
            If .HURT = True Then StretchBlt memBuffer, .PX + .Xpos(W), .PY + .Ypos(W), .Widthp(W), .Heightp(W), .memHurt, .Xp(W), .Yp(W), .Widthp(W), .Heightp(W), &HEE0086 ' SRCPAINT
            StretchBlt memBuffer, .PX + .Xpos(W), .PY + .Ypos(W), .Widthp(W), .Heightp(W), .memSpr, .Xp(W), .Yp(W), .Widthp(W), .Heightp(W), &HEE0086 ' SRCPAINT
        End If
    'Draw Alucard's effects
        If .DrwWep = True And .pLeft = True Then
            StretchBlt memBuffer, .PX + .SXP + .WepW(.WEP), .PY + .WepYpos(.WEP), -.WepW(.WEP), .WepH(.WEP), .memWepMsk, .WepX(.WEP), .WepY(.WEP), .WepW(.WEP), .WepH(.WEP), &H8800C6 ' SRCAND
            StretchBlt memBuffer, .PX + .SXP + .WepW(.WEP), .PY + .WepYpos(.WEP), -.WepW(.WEP), .WepH(.WEP), .memWepSpr, .WepX(.WEP), .WepY(.WEP), .WepW(.WEP), .WepH(.WEP), &HEE0086 ' SRCPAINT
        ElseIf .DrwWep = True And .pLeft = False Then
            StretchBlt memBuffer, .PX + .WepXpos(.WEP), .PY + .WepYpos(.WEP), .WepW(.WEP), .WepH(.WEP), .memWepMsk, .WepX(.WEP), .WepY(.WEP), .WepW(.WEP), .WepH(.WEP), &H8800C6 ' SRCAND
            StretchBlt memBuffer, .PX + .WepXpos(.WEP), .PY + .WepYpos(.WEP), .WepW(.WEP), .WepH(.WEP), .memWepSpr, .WepX(.WEP), .WepY(.WEP), .WepW(.WEP), .WepH(.WEP), &HEE0086 ' SRCPAINT
        End If
    End With
    'Draw Items
    For Obj = 1 To TotlItm
        BitBlt memBuffer, BgX + IXpos(Obj), BgY + IYpos(Obj), IWp(ItmType(Obj)), IHp(ItmType(Obj)), memItmMsk, IXp(ItmType(Obj)), IYp(ItmType(Obj)), &H8800C6 ' SRCAND
        BitBlt memBuffer, BgX + IXpos(Obj), BgY + IYpos(Obj), IWp(ItmType(Obj)), IHp(ItmType(Obj)), memItmSpr, IXp(ItmType(Obj)), IYp(ItmType(Obj)), &HEE0086 ' SRCPAINT
    Next
    'Draw Projectiles & Sub Weapons
    For Obj = 1 To TotlSub
        With Enemy(Obj)
            W = .SWEP
            If .wLeft = True Then
                StretchBlt memBuffer, BgX + .WX + .WXP + EWepW(W), BgY + .WY + EWepYpos(W), -EWepW(W), EWepH(W), memSubMsk, EWepX(W), EWepY(W), EWepW(W), EWepH(W), &H8800C6 ' SRCAND
                StretchBlt memBuffer, BgX + .WX + .WXP + EWepW(W), BgY + .WY + EWepYpos(W), -EWepW(W), EWepH(W), memSubSpr, EWepX(W), EWepY(W), EWepW(W), EWepH(W), &HEE0086 ' SRCPAINT
            ElseIf .wLeft = False Then
                BitBlt memBuffer, BgX + .WX + EWepXpos(W), BgY + .WY + EWepYpos(W), EWepW(W), EWepH(W), memSubMsk, EWepX(W), EWepY(W), &H8800C6 ' SRCAND
                BitBlt memBuffer, BgX + .WX + EWepXpos(W), BgY + .WY + EWepYpos(W), EWepW(W), EWepH(W), memSubSpr, EWepX(W), EWepY(W), &HEE0086 ' SRCPAINT
            End If
        End With
    Next
    'Draw Effects
    CH = 0
    If pblnDrawSplash(CH) = True Then
        BitBlt memBuffer, EfctX(CH) + BgX + PsXpos(AniEfct(CH)), EfctY(CH) + BgY + PsYpos(AniEfct(CH)), PsWp(AniEfct(CH)), PsHp(AniEfct(CH)), memBgPsMsk, PsXp(AniEfct(CH)), PsYp(AniEfct(CH)), &H8800C6 ' SRCAND
        BitBlt memBuffer, EfctX(CH) + BgX + PsXpos(AniEfct(CH)), EfctY(CH) + BgY + PsYpos(AniEfct(CH)), PsWp(AniEfct(CH)), PsHp(AniEfct(CH)), memBgPsSpr, PsXp(AniEfct(CH)), PsYp(AniEfct(CH)), &HEE0086 ' SRCPAINT
    End If
    'Draw Un-Passable background objects
    For Obj = 1 To TotlUp
        BitBlt memBuffer, BgX + UpXpos(Obj), BgY + UpYpos(Obj), UpWp(Obj), UpHp(Obj), memBgUpMsk, UpXp(Obj), UpYp(Obj), &H8800C6 ' SRCAND
        BitBlt memBuffer, BgX + UpXpos(Obj), BgY + UpYpos(Obj), UpWp(Obj), UpHp(Obj), memBgUpSpr, UpXp(Obj), UpYp(Obj), &HEE0086 ' SRCPAINT
    Next
    'Draw Headsup Display
    For Obj = 1 To 3
        If Obj = 3 Then DW = HWidth(Head(Obj)) * (Player.HP / Player.MaxHP) Else DW = HWidth(Head(Obj)) ' Bar width
        BitBlt memBuffer, HXpos(Head(Obj)), HYpos(Head(Obj)), DW, HHeight(Head(Obj)), memHeadMsk, HXp(Head(Obj)), HYp(Head(Obj)), &H8800C6 ' SRCAND
        BitBlt memBuffer, HXpos(Head(Obj)), HYpos(Head(Obj)), DW, HHeight(Head(Obj)), memHeadSpr, HXp(Head(Obj)), HYp(Head(Obj)), &HEE0086 ' SRCPAINT
    Next
    'Draw Headsup Numbers
    H = Player.HP
    Select Case Len(H)
        Case 1: P = 39
        Case 2: P = 35
        Case 3: P = 31
        Case Is >= 4: P = 28
    End Select
    For Obj = 4 To HeadCount
        BitBlt memBuffer, P, 36, HWidth(Head(Obj)), HHeight(Head(Obj)), memHeadMsk, HXp(Head(Obj)), HYp(Head(Obj)), &H8800C6 ' SRCAND
        BitBlt memBuffer, P, 36, HWidth(Head(Obj)), HHeight(Head(Obj)), memHeadSpr, HXp(Head(Obj)), HYp(Head(Obj)), &HEE0086 ' SRCPAINT
        P = P + 7
    Next
    ' Show Items - Uses mblnItems in keydown event
    If mblnItems = True Then
        For ItemLoop = 1 To 10
            Me.CurrentX = 0
            Me.CurrentY = 0
            Me.Font.Name = "Verdana"
            Me.Font.Size = 7
            Me.ForeColor = vbGreen
            For Obj = 1 To 20
                Me.Print "Type " & Obj; " Count " & Inven(Obj)
            Next
        Next
    End If
    BitBlt memScreen, 0, 0, SCREENWIDTH, SCREENHEIGHT, memBuffer, 0, 0, &HCC0020 ' SRCCOPY
    ' Draw Collision Box
    'Me.Line (BL(0) + PX(0), BT(0) + PY(0))-(BW(0) + PX(0), BH(0) + PY(0)), vbBlue, B
    
    'View Array
    'For intType = 1 To 2
    '    For intAnim = 1 To 20
    '        Debug.Print intType; intAnim; EXpos(intAnim, intType); EYpos(intAnim, intType); EXp(intAnim, intType); EYp(intAnim, intType); EWidthp(intAnim, intType); EHeightp(intAnim, intType)
    '    Next
    'Next
    
    
    'Show Memory Usage
    ''GlobalMemoryStatus memInfo
    ''R1 = memInfo.dwTotalPhys
    ''R2 = memInfo.dwAvailPhys
    ''With Me
    ''    .Font.Name = "Verdana"
    ''    .ForeColor = 65280
    ''    .FontSize = 20
    ''    .CurrentX = 10
    ''    .CurrentY = 10
    ''End With
    ''Me.Print R1 - R2
    'T2 = GetTickCount()
    'Debug.Print T2 - T1; "Draw Screen"
    'Debug.Print TotlCand
    'Debug.Print totlen
    'Debug.Print TotlUp
    'Debug.Print HeadCount
    'Debug.Print PX(0)
End Sub

'Private Sub FlipScreen()
'        BitBlt memScreen, 0, 0, SCREENWIDTH, SCREENHEIGHT, memBuffer, 0, 0, &HCC0020 ' SRCCOPY
'End Sub


'/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\
' MAIN ENGINE
' -------------------------------------------------------------------
' KEYDOWN AND KEYUP
'\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    'If KeyCode = vbKeyRight Or KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then
    '    mblnAttack = False 'Gives new movement priority over attacks
    'End If
    Select Case KeyCode
        Case vbKey0: DELAY = 35
        Case vbKey1: DELAY = 1000
        Case vbKey2: DELAY = 2000
        Case vbKey3: DELAY = 3000
        Case vbKey4: DELAY = 4000
        Case vbKey5: DELAY = 5000
        Case vbKey9: DELAY = 0
        Case vbKeyEscape
            GameOver = True 'mblnEndProgram = True
        Case vbKeyRight
            mblnRight = True
            mblnLeft = False
        Case vbKeyLeft
            mblnLeft = True
            mblnRight = False
        Case vbKeyDown
            mblnDown = True
        Case vbKeyZ
            If mblnAttack = False Then mblnUp = True
        Case vbKeyX
            mblnAttack = True
            mblnUp = False
        Case vbKeyM ' Print items to screen
            If mblnItems = True Then
                mblnItems = False
            Else
                mblnItems = True
            End If
    End Select
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    With Player
        Select Case KeyCode
            Case vbKeyRight
                If mblnRight = True Then
                    mblnRight = False
                    ResetTurn
                    If .pLeft = False Then
                        If .STA <> JUMPRIGHT And .STA <> FALLRIGHT And .STA <> LANDRIGHT And .STA <> FALLMOVERIGHT And .STA <> FALLLEFT And .STA <> LANDLEFT And .STA <> FALLMOVELEFT And .STA <> DASHRIGHT And .STA <> ATTACKDOWNRIGHT And .STA <> DUCKRIGHT And .STA <> JUMPATTACKRIGHT And .STA <> FALLATTACKRIGHT Then
                            .STA = STOPRIGHT
                        End If
                    End If
                End If
            Case vbKeyLeft
                If mblnLeft = True Then
                    mblnLeft = False
                    ResetTurn
                    If .pLeft = True Then
                        If .STA <> JUMPLEFT And .STA <> FALLRIGHT And .STA <> LANDRIGHT And .STA <> FALLMOVERIGHT And .STA <> FALLLEFT And .STA <> LANDLEFT And .STA <> FALLMOVELEFT And .STA <> DASHLEFT And .STA <> ATTACKDOWNLEFT And .STA <> DUCKLEFT And .STA <> JUMPATTACKLEFT And .STA <> FALLATTACKLEFT Then
                            .STA = STOPLEFT
                        End If
                    End If
                End If
            Case vbKeyDown
                mblnDown = False
            Case vbKeyZ
                mblnUp = False
                .JT = 0
                If .STA = JUMPRIGHT And .pLeft = False Then
                    .STA = FALLRIGHT
                ElseIf .STA = JUMPLEFT And .pLeft = True Then
                    .STA = FALLLEFT
                End If
        End Select
    End With
End Sub


'/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\
' KEYDOWN AND KEYUP
' -------------------------------------------------------------------
' ALUCARD MOVEMENT
'\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/


Private Function ActionWalk(W As Long) As Long
'Dash 26 to 40
'Walk 41 to 56
    With Player
        .Grav = 0
        If W < 25 Or W > 56 Then W = 25
        If .pLeft = True Then
            MoveLeft DISTANCE, SPACE
        Else
            MoveRight DISTANCE, SPACE
        End If
        If W >= 25 And W < 41 Then
            .STA = IIf(.pLeft = True, DASHLEFT, DASHRIGHT)
            W = W + 1
        ElseIf W >= 41 And W < 56 Then
            .STA = IIf(.pLeft = True, WALKLEFT, WALKRIGHT)
            W = W + 1
        ElseIf W = 56 Then
            W = 41
        End If
        'Debug.Print "Walk"; W; pLeft(CH)
    End With
End Function

Private Function ActionTurnAround(W As Long, NewL As Boolean) As Long
'Turn 57 to 66
    With Player
        .Grav = 0
        If .pLeft <> NewL Or (W < 56 Or W > 66) Then W = 57 ' Direction is checked to restart turn when changing direction while turning.
        .pLeft = NewL ' Set New Direction
        If .pLeft = True Then
            MoveLeft DISTANCE, SPACE
        Else
            MoveRight DISTANCE, SPACE
        End If
        If .AniDelay >= 1 Then
           .AniDelay = 0
            If W >= 56 And W <= 65 Then
                .STA = IIf(.pLeft = True, TURNLEFT, TURNRIGHT)
                W = W + 1
            ElseIf W = 66 Then
                W = 49 ' Middle Of Walking
                .STA = IIf(.pLeft = True, WALKLEFT, WALKRIGHT)
            End If
        Else
            .AniDelay = .AniDelay + 1
        End If
        'Debug.Print "Turn"; W; pLeft(CH)
    End With
End Function

Private Function ActionStop(W As Long) As Long
'Stop 67 to 79
    With Player
        .Grav = 0
        If W < 66 Or W > 79 Then W = 66
        If W >= 66 And W <= 77 Then
            W = W + 1
        ElseIf W = 78 Then
            .STA = IIf(.pLeft = True, STANDLEFT, STANDRIGHT)
            W = W + 1
        End If
        'Debug.Print "Stop"; W; pLeft(CH)
    End With
End Function

Private Function ActionDuck(W As Long) As Long
'Duck 2 to 14
    With Player
        .Grav = 0
        If W < 1 Or W > 15 Then W = 1
        If W >= 1 And W < 13 Then
            .STA = IIf(.pLeft = True, DUCKLEFT, DUCKRIGHT)
            W = W + 1
        End If
        'Debug.Print "Duck"; W; pLeft(CH); STA(CH)
    End With
End Function

Private Function ActionFall(W As Long) As Long
'Fall 108 to 116
    With Player
        If W < 108 Or W > 116 Then W = 108
        .STA = IIf(.pLeft = True, FALLLEFT, FALLRIGHT)
        If W >= 107 And W < 116 Then
            W = W + 1
        End If
        'Debug.Print "Fall"; W; pLeft(CH); STA(CH)
    End With
End Function

Private Function ActionLand(W As Long) As Long
'Land 118 to 122
    With Player
        mblnAttack = False
        .Grav = 0
        If W < 117 Or W > 122 Then W = 117
        If W >= 117 And W < 122 Then
            .STA = IIf(.pLeft = True, LANDLEFT, LANDRIGHT)
            W = W + 1
        ElseIf W >= 122 Then
            .STA = IIf(.pLeft = True, STANDLEFT, STANDRIGHT)
        End If
        'Debug.Print "Land"; W; pLeft(CH)
    End With
End Function

Private Function ActionHardLand(W As Long) As Long
'Hard Land 195 to 203
    With Player
        mblnAttack = False
        If W < 194 Or W > 203 Then W = 194
        If W >= 194 And W < 203 Then
            .STA = IIf(.pLeft = True, LANDLEFT, LANDRIGHT)
            W = W + 1
        ElseIf W >= 203 Then
            .Grav = 0
            .STA = IIf(.pLeft = True, STANDUPLEFT, STANDUPRIGHT)
        End If
        'Debug.Print "Hard Land"; W; pLeft(CH)
    End With
End Function

Private Function ActionStand(W As Long) As Long
'Stand 14 to 18
    With Player
        .Grav = 0
        If W < 13 Or W > 18 Then W = 13
        If W >= 13 And W < 18 Then
            .STA = IIf(.pLeft = True, STANDUPLEFT, STANDUPRIGHT)
            W = W + 1
        ElseIf W >= 18 Then
            .STA = IIf(.pLeft = True, STANDLEFT, STANDRIGHT)
        End If
        'Debug.Print "Stand"; W; pLeft(CH)
    End With
End Function

Private Function ActionHardHurt(W As Long) As Long
Dim DMG As Long
'178 to 180
    With Player
        mblnAttack = False
        .Grav = 20
        Randomize
        ' Start Animation & set pLeft(CH)
        If W < 175 Or W > 179 Then
            W = 175
            Select Case 2 * Rnd ' Pick Right or Left
                Case Is <= 1: .pLeft = True
                Case Is >= 2: .pLeft = False
            End Select
        End If
        ' Move Character
        If .pLeft = True Then
            MoveUp 5, 6
            MoveRight 3, 4
        Else
            MoveUp 5, 6
            MoveLeft 3, 4
        End If
        ' Change Frame
        If .AniDelay >= Rnd * 5 Then
            .AniDelay = 0
            If W >= 175 And W <= 178 Then
                .STA = IIf(.pLeft = True, HURTLEFT, HURTRIGHT)
                W = W + 1
            ElseIf W >= 179 Then
                .STA = IIf(.pLeft = True, FALLLEFT, FALLRIGHT)
                W = W + 1
                .HURT = False
            End If
            ' Game Over
            If .HP <= 0 Then GameOver = True
        Else
            .AniDelay = .AniDelay + 1
        End If
        ' Stop When A Surface Is Hit
        If .WallR = True Or .WallL = True Or .WallD = True Then
            .STA = IIf(.pLeft = True, FALLLEFT, FALLRIGHT)
            .HURT = False
        End If
        'Debug.Print "Hard Hurt"; W; pLeft(CH)
    End With
End Function

Private Function ActionAirHurt(W As Long) As Long
Dim H As Long, DMG As Long
'161 to 163
    With Player
        mblnAttack = False
        .Grav = 20
        Randomize
        If W < 161 Or W > 163 Then
            DMG = Enemy(.HrtingCh).AP - (.DP \ Enemy(.HrtingCh).AP) ' Calculate Damage
            If DMG <= 0 Then DMG = 1
            .HP = .HP - DMG
            If .HP < 0 Then .HP = 0 ' No negative numbers
            Select Case 2 * Rnd ' Pick Right or Left
                Case Is <= 1: .pLeft = True
                Case Is >= 2: .pLeft = False
            End Select
        End If
        If .pLeft = True Then
            MoveUp 5, 6
            MoveRight 5, 6
        Else
            MoveUp 5, 6
            MoveLeft 5, 6
        End If
        If W < 161 Or W > 163 Then
            .STA = IIf(.pLeft = True, HURTLEFT, HURTRIGHT)
            H = Rnd * 3
            Select Case H
                Case 1: W = 161
                Case 2: W = 162
                Case 3: W = 163
            End Select
        End If
        If .AniDelay >= 6 Then
            .AniDelay = 0
            .STA = IIf(.pLeft = True, STANDLEFT, STANDRIGHT)
            .HURT = False
        Else
            .AniDelay = .AniDelay + 1
        End If
        ' Stop When A Surface Is Hit
        If .WallR = True Or .WallL = True Or .WallD = True Then
            .STA = IIf(.pLeft = True, FALLLEFT, FALLRIGHT)
            .HURT = False
        End If
        'Debug.Print "Air Hurt"; W; pLeft(CH)
    End With
End Function

Private Function ActionHurt(W As Long) As Long
Dim R As Integer, H As Long, DMG As Long
'164 to 168
    With Player
        mblnAttack = False
        .Grav = 20
        Randomize
        If W < 164 Or W > 168 Then
            DMG = Enemy(.HrtingCh).AP - (.DP \ Enemy(.HrtingCh).AP) ' Calculate Damage
            If DMG <= 0 Then DMG = 1
            .HP = .HP - DMG
            If .HP < 0 Then .HP = 0 ' No negative numbers
            Select Case 2 * Rnd ' Pick Right or Left
                Case Is <= 1: .pLeft = True
                Case Is >= 2: .pLeft = False
            End Select
        End If
        R = Rnd * 3
        If .pLeft = True Then
            MoveUp R, R + 1
            MoveRight 5, 6
        Else
            MoveUp R, R + 1
            MoveLeft 5, 6
        End If
        If W < 164 Or W > 168 Then
            .STA = IIf(.pLeft = True, HURTLEFT, HURTRIGHT)
            H = Rnd * 5
            Select Case H
                Case 1: W = 164
                Case 2: W = 165
                Case 3: W = 166
                Case 4: W = 167
                Case 5: W = 168
            End Select
        End If
        If .AniDelay >= 6 Then
            .AniDelay = 0
            .STA = IIf(.pLeft = True, STANDLEFT, STANDRIGHT)
            .HURT = False
        Else
            .AniDelay = .AniDelay + 1
        End If
        'Debug.Print "Hurt"; W; pLeft(CH)
    End With
End Function

Private Function ActionAirMove(W As Long) As Long
'Fall 110 to 118
    With Player
        If W < 110 Or W > 118 Then W = 110
        .STA = IIf(.pLeft = True, FALLLEFT, FALLRIGHT)
        If .pLeft = True Then
            MoveLeft DISTANCE, SPACE
        Else
            MoveRight DISTANCE, SPACE
        End If
        If W >= 107 And W < 116 Then W = W + 1
        'Debug.Print "Air Move"; W; pLeft(CH)
    End With
End Function

Private Function ActionJump(W As Long) As Long
'Jump 102 to 107
    With Player
        .Grav = 0
        ' Reset Frame & Jump
        If W < 101 Or W > 107 Then
            W = 102
        End If
        ' Move Up
        MoveUp CInt(.JD), CInt(.JS)
        If .JD - 0.2 > 1 Then .JD = .JD - 0.2
        If .JS - 3.2 > 1 Then .JS = .JS - 0.2
        ' Move Direction Facing
        If .pLeft = True Then
            MoveLeft DISTANCE, SPACE
        Else
            MoveRight DISTANCE, SPACE
        End If
        ' Change Frame
        If .AniDelay >= 1 Then
            .AniDelay = 0
            If W >= 101 And W < 107 Then
                .STA = IIf(.pLeft = True, JUMPLEFT, JUMPRIGHT)
                W = W + 1
            ElseIf W = 107 Then
                W = 104
            End If
        Else
            .STA = IIf(.pLeft = True, JUMPLEFT, JUMPRIGHT)
            .AniDelay = .AniDelay + 1
        End If
        ' Record Distance Jumped
        .JT = .JT + CInt(.JD)
        If .JT > JUMPHEIGHT Then
            .STA = IIf(.pLeft = True, FALLLEFT, FALLRIGHT)
            .JT = 0
        End If
        'Debug.Print "Jump"; W; pLeft(CH); STA(CH)
    End With
End Function

Private Function ActionJumpUp(W As Long) As Long
'Jump Up 95 to 101
    With Player
        .Grav = 0
        If W < 95 Or W > 101 Then
            W = 95
        End If
        ' Move Up
        MoveUp CInt(.JD), CInt(.JS)
        If .JD - 0.2 > 1 Then .JD = .JD - 0.2
        If .JS - 3.2 > 1 Then .JS = .JS - 0.2
        If .AniDelay >= 1 Then
            .AniDelay = 0
            If W >= 95 And W < 101 Then
                .STA = IIf(.pLeft = True, JUMPLEFT, JUMPRIGHT)
                W = W + 1
            End If
        Else
            .STA = IIf(.pLeft = True, JUMPLEFT, JUMPRIGHT)
            .AniDelay = .AniDelay + 1
        End If
        .JT = .JT + CInt(.JD)
        If .JT > JUMPHEIGHT Then
            .STA = IIf(.pLeft = True, FALLLEFT, FALLRIGHT)
            .JT = 0
        End If
        'Debug.Print "Jump Up"; W; pLeft(CH); STA(CH)
    End With
End Function

Private Function ActionSword(W As Long) As Long 'Self terminates because last mintAnim is larger than the setStart
'Sword 181 to 190
    With Player
        .DrwWep = True
        .Grav = 0
        If W < 180 Or W > 190 Then
            W = 180
            If blnSndFX(1) = True Then SndEft(1).Play DSBPLAY_DEFAULT
        End If
        If W >= 180 And W <= 188 Then
            .STA = IIf(.pLeft = True, ATTACKLEFT, ATTACKRIGHT)
            W = W + 1
        ElseIf W >= 189 Then
            .STA = IIf(.pLeft = True, STANDLEFT, STANDRIGHT)
            W = W + 1
            mblnAttack = False
        End If
        If W <= 197 And .STA = ATTACKRIGHT Or .STA = ATTACKLEFT Then
            Sword 2, 9, 1
        Else
            .DrwWep = False ' Animate Sword Until This Frame
        End If
        'Debug.Print "Sword"; W; pLeft(CH)
    End With
End Function

Private Function ActionSwordDown(W As Long) As Long
'Sword Down 192 to 203
    With Player
        .DrwWep = True
        .Grav = 0
        If W < 191 Or W > 203 Then
            W = 191
            If blnSndFX(2) = True Then SndEft(2).Play DSBPLAY_DEFAULT
        End If
        If W >= 191 And W <= 202 Then
            .STA = IIf(.pLeft = True, ATTACKDOWNLEFT, ATTACKDOWNRIGHT)
            W = W + 1
        ElseIf W >= 203 Then
            .STA = IIf(.pLeft = True, DUCKLEFT, DUCKRIGHT)
            W = 14
            mblnAttack = False
        End If
        If W = 193 Then W = 194 ' Skip this frame
        If W <= 199 And W <> 14 Then
            Sword 10, 16, 9
        Else
            .DrwWep = False ' Animate Sword Until This Frame
        End If
        'Debug.Print "Sword Down"; W; pLeft(CH)
    End With
End Function

Private Function ActionSwordOut(W As Long) As Long
'Sword Down 192 to 203
    With Player
        .DrwWep = True
        .Grav = 0
        If W < 191 Or W > 203 Then W = 191
        If W >= 191 And W <= 202 Then
            .STA = IIf(.pLeft = True, ATTACKOUTLEFT, ATTACKOUTRIGHT)
            W = W + 1
        ElseIf W >= 203 Then
            .STA = IIf(.pLeft = True, DUCKLEFT, DUCKRIGHT)
            W = 14
            mblnAttack = False
        End If
        If W = 194 Then W = 195 ' Skip this frame
        If W <= 199 And W <> 14 Then
            Sword 17, 23, 16
        Else
            .DrwWep = False ' Animate Sword Until This Frame
        End If
        'Debug.Print "Sword Out"; W; pLeft(CH)
    End With
End Function

Private Function ActionJumpSwordOut(W As Long) As Long
'Air Sword Out 204 to 207
    With Player
        .DrwWep = True
        .Grav = 0
        If W < 204 Or W > 207 Then
            W = 203
        End If
        ' Move Up
        MoveUp CInt(.JD), CInt(.JS)
        If .JD - 0.2 > 1 Then .JD = .JD - 0.2
        If .JS - 3.2 > 1 Then .JS = .JS - 0.2
        ' Move If Desired
        If .pLeft = True And mblnLeft = True Then
            MoveLeft DISTANCE, SPACE
        ElseIf .pLeft = False And mblnRight = True Then
            MoveRight DISTANCE, SPACE
        End If
        If W >= 203 And W <= 206 Then
            .STA = IIf(.pLeft = True, JUMPATTACKLEFT, JUMPATTACKRIGHT)
            W = W + 1
            .JT = .JT + CInt(.JD)
        ElseIf W >= 207 Then
            .STA = IIf(.pLeft = True, FALLLEFT, FALLRIGHT)
            W = 108
            .JT = 0 ' Reset Jumping
            mblnAttack = False
        End If
        If W <= 207 And W <> 108 Then
            Sword 24, 27, 23
        Else
            .DrwWep = False ' Animate Sword Until This Frame
        End If
        'Debug.Print "Sword Air Out"; W; pLeft(CH); STA(CH)
    End With
End Function

Private Function ActionFallSwordDown(W As Long) As Long
'Air Sword Down 208 to 212
    With Player
        .DrwWep = True
        .Grav = 0
        If W < 208 Or W > 212 Then
            W = 207
        Else ' Don't Move On First Frame Because Gravity Takes Care Of That
            If .pLeft = True And mblnLeft = True Then ' Left Only If Pushing Left
                MoveLeft DISTANCE, SPACE
            ElseIf .pLeft = False And mblnRight = True Then ' Right Only If Pushing Right
                MoveRight DISTANCE, SPACE
            End If
        End If
        If W >= 207 And W <= 211 Then
            .STA = IIf(.pLeft = True, FALLATTACKLEFTDOWN, FALLATTACKRIGHTDOWN)
            W = W + 1
        ElseIf W >= 212 Then
            'Delay Used To Keep Last Frame For More Than One Loop
            If .AniDelay >= 2 Then
                .AniDelay = 0
                .STA = IIf(.pLeft = True, FALLLEFT, FALLRIGHT)
                W = 114
                .JT = 0 ' Reset Jumping
                mblnAttack = False
            Else
                .AniDelay = .AniDelay + 1
            End If
        End If
        If W = 209 Then W = 210 ' Skip this frame
        If W <= 213 And W <> 114 Then
            Sword 29, 33, 28
        Else
            .DrwWep = False ' Animate Sword Until This Frame
        End If
    End With
        'Debug.Print "Sword Air Down"; W; pLeft(CH); STA(CH)
End Function

Private Function ActionFallSwordOut(W As Long) As Long
'Air Sword Down 208 to 212
    With Player
        .DrwWep = True
        .Grav = 0
        If W < 208 Or W > 212 Then
            W = 207
        Else ' Don't Move On First Frame Because Gravity Takes Care Of That
            If .pLeft = True And mblnLeft = True Then ' Left Only If Pushing Left
                MoveLeft DISTANCE, SPACE
            ElseIf .pLeft = False And mblnRight = True Then ' Right Only If Pushing Right
                MoveRight DISTANCE, SPACE
            End If
        End If
        If W >= 207 And W <= 211 Then
            .STA = IIf(.pLeft = True, FALLATTACKLEFT, FALLATTACKRIGHT)
            W = W + 1
        ElseIf W >= 212 Then
            'Delay Used To Keep Last Frame For More Than One Loop
            If .AniDelay >= 2 Then
                .AniDelay = 0
                .STA = IIf(.pLeft = True, FALLLEFT, FALLRIGHT)
                W = 114
                .JT = 0 ' Reset Jumping
                mblnAttack = False
            Else
                .AniDelay = .AniDelay + 1
            End If
        End If
        If W = 210 Then W = 211 ' Skip this frame
        If W <= 212 And W <> 114 Then
            Sword 34, 38, 33
        Else
            .DrwWep = False ' Animate Sword Until This Frame
        End If
        'Debug.Print "Sword Air Out"; W; pLeft(CH); STA(CH)
    End With
End Function

Private Function ActionBob(W As Long) As Long 'Loops because last mintAnim sets it to the first
'Bob 20 to 24
    With Player
        .Grav = 0
        If W < 19 Or W > 24 Then W = 19
        If .AniDelay >= 3 Then
            .AniDelay = 0
            If W >= 19 And W < 24 Then
                .STA = IIf(.pLeft = True, STANDLEFT, STANDRIGHT)
                W = W + 1
            ElseIf W = 24 Then
                W = 20
            End If
        Else
            .AniDelay = .AniDelay + 1
        End If
        'Debug.Print "Bob"; W; pLeft(CH)
    End With
End Function

Private Function ActionTeeter(W As Long) As Long
'Teeter 86 to 94
    With Player
        .Grav = 0
        If W < 85 Or W > 94 Then W = 85
        If .AniDelay >= 2 Then
            .AniDelay = 0
            If W >= 85 And W < 94 Then
                .STA = IIf(.pLeft = True, STOPLEFT, STOPRIGHT)
                W = W + 1
            End If
        Else
            .AniDelay = .AniDelay + 1
        End If
        'Debug.Print "Teeter"; W; pLeft(CH)
    End With
End Function

Private Function ActionStill(W As Long) As Long
'Teeter 86 to 94
'Still 89 to 91
    With Player
        .Grav = 0
        If W < 88 Or W > 91 Then W = 89
        If .AniDelay >= 2 Then
            .AniDelay = 0
            If W >= 88 And W < 91 Then
                .STA = IIf(.pLeft = True, STANDLEFT, STANDRIGHT)
                W = W + 1
            End If
        Else
            .AniDelay = .AniDelay + 1
        End If
        'Debug.Print "Still"; W; pLeft(CH)
    End With
End Function


'/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\
' ALUCARD MOVEMENT
' -------------------------------------------------------------------
' MOVEMENT
'\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/


Private Function MoveAlucard(S As Long, W As Long) As Long
Dim retVal As Boolean
    With Player
        .DrwWep = False ' resets weapon
        ' Death Animation
        If .HP <= 0 Then
            ActionHardHurt W
        ' Hurt Animations
        ElseIf .HURT = True Then
            If (S = HURTRIGHT Or S = DUCKRIGHT Or S = STANDRIGHT Or S = WALKRIGHT Or S = DASHRIGHT Or S = ATTACKRIGHT Or S = ATTACKOUTRIGHT Or S = ATTACKDOWNRIGHT Or S = STANDUPRIGHT Or S = TURNRIGHT Or S = LANDRIGHT Or S = STOPRIGHT) Then
                ActionHurt W
            ElseIf (S = HURTLEFT Or S = DUCKLEFT Or S = STANDLEFT Or S = WALKLEFT Or S = DASHLEFT Or S = ATTACKLEFT Or S = ATTACKOUTLEFT Or S = ATTACKDOWNLEFT Or S = STANDUPLEFT Or S = TURNLEFT Or S = LANDLEFT Or S = STOPLEFT) Then
                ActionHurt W
            Else
                ActionAirHurt W
            End If
        ' Jumping Animations
        ElseIf mblnUp = True Then
            If mblnLeft = True And .WallU = False And (S = JUMPLEFT Or S = JUMPRIGHT Or S = STANDLEFT Or S = DASHLEFT Or S = WALKLEFT Or S = TURNLEFT Or S = STANDUPLEFT Or S = DUCKLEFT Or S = STOPLEFT Or S = LANDLEFT) Then
                .pLeft = True
                ActionJump W
            ElseIf mblnRight = True And .WallU = False And (S = JUMPRIGHT Or S = JUMPLEFT Or S = STANDRIGHT Or S = DASHRIGHT Or S = WALKRIGHT Or S = TURNRIGHT Or S = STANDUPRIGHT Or S = DUCKRIGHT Or S = STOPRIGHT Or S = LANDRIGHT) Then
                .pLeft = False
                ActionJump W
            ElseIf mblnRight = False And mblnLeft = False And .WallU = False And (S = JUMPRIGHT Or S = STANDRIGHT Or S = DASHRIGHT Or S = WALKRIGHT Or S = TURNRIGHT Or S = STANDUPRIGHT Or S = DUCKRIGHT Or S = STOPRIGHT Or S = LANDRIGHT) Then
                .pLeft = False
                ActionJumpUp W
            ElseIf mblnRight = False And mblnLeft = False And .WallU = False And (S = JUMPLEFT Or S = STANDLEFT Or S = DASHLEFT Or S = WALKLEFT Or S = TURNLEFT Or S = STANDUPLEFT Or S = DUCKLEFT Or S = STOPLEFT Or S = LANDLEFT) Then
                .pLeft = True
                ActionJumpUp W
            Else
                AlucardGravity S, W
            End If
        ' Walking Right Animations
        ElseIf mblnRight = True And mblnDown = False And mblnAttack = False Then
            AlucardGravity S, W
            If (S = STANDUPLEFT Or S = LANDLEFT Or S = STOPLEFT Or S = TURNRIGHT Or S = STANDLEFT Or S = DUCKLEFT Or S = DASHLEFT Or S = WALKLEFT Or S = TURNLEFT) Then
                ActionTurnAround W, False
            ElseIf (S = STANDUPRIGHT Or S = LANDRIGHT Or S = STOPRIGHT Or S = WALKRIGHT Or S = DASHRIGHT Or S = STANDRIGHT Or S = DUCKRIGHT) Then
                ActionWalk W
            End If
        ' Walking Left Animations
        ElseIf mblnLeft = True And mblnDown = False And mblnAttack = False Then
            AlucardGravity S, W
            If (S = STANDUPRIGHT Or S = LANDRIGHT Or S = STOPRIGHT Or S = TURNLEFT Or S = STANDRIGHT Or S = DUCKRIGHT Or S = DASHRIGHT Or S = WALKRIGHT Or S = TURNRIGHT) Then
                ActionTurnAround W, True
            ElseIf (S = STANDUPLEFT Or S = LANDLEFT Or S = STOPLEFT Or S = WALKLEFT Or S = DASHLEFT Or S = STANDLEFT Or S = DUCKLEFT) Then
                ActionWalk W
            End If
        ' Ducking & Sinking Animations
        ElseIf mblnDown = True And mblnAttack = False Then
            AlucardGravity S, W
            If mblnUp = True And .Soft = True And (S = DUCKRIGHT) Then
                .mblnSink = True
                .pLeft = False
                ActionFall W
                MoveDown CInt(.FD), CInt(.FS)
            ElseIf mblnUp = True And .Soft = True And (S = DUCKLEFT) Then
                .mblnSink = True
                .pLeft = True
                ActionFall W
                MoveDown CInt(.FD), CInt(.FS)
            ElseIf (S = STANDUPRIGHT Or S = LANDRIGHT Or S = DUCKRIGHT Or S = STOPRIGHT Or S = STANDRIGHT Or S = DASHRIGHT Or S = WALKRIGHT Or S = TURNRIGHT) Then
                ActionDuck W
            ElseIf (S = STANDUPLEFT Or S = LANDLEFT Or S = DUCKLEFT Or S = STOPLEFT Or S = STANDLEFT Or S = DASHLEFT Or S = WALKLEFT Or S = TURNLEFT) Then
                ActionDuck W
            End If
        ' Attacking Animations
        ElseIf mblnAttack = True Then
            retVal = AlucardGravity(S, W)
            If retVal = False Then ' No Landing Animation
                If .pLeft = True And mblnUp = True And .WallU = False And (S = JUMPATTACKLEFT Or S = JUMPLEFT Or S = STANDLEFT Or S = DASHLEFT Or S = WALKLEFT Or S = TURNLEFT Or S = STANDUPLEFT Or S = DUCKLEFT Or S = STOPLEFT Or S = LANDLEFT) Then
                    ActionJumpSwordOut W
                ElseIf .pLeft = False And mblnUp = True And .WallU = False And (S = JUMPATTACKRIGHT Or S = JUMPRIGHT Or S = STANDRIGHT Or S = DASHRIGHT Or S = WALKRIGHT Or S = TURNRIGHT Or S = STANDUPRIGHT Or S = DUCKRIGHT Or S = STOPRIGHT Or S = LANDRIGHT) Then
                    ActionJumpSwordOut W
                ElseIf .pLeft = True And .WallD = False And .WallU = False And (S = FALLATTACKLEFTDOWN Or S = FALLATTACKLEFT Or S = FALLLEFT Or FALLMOVELEFT) Then
                    If (mblnDown = True Or S = FALLATTACKLEFTDOWN) And S <> FALLATTACKLEFT Then
                        .pLeft = True
                        ActionFallSwordDown W
                    ElseIf mblnDown <> True Or S = FALLATTACKLEFT Then
                        .pLeft = True
                        ActionFallSwordOut W
                    End If
                ElseIf .pLeft = False And .WallD = False And .WallU = False And (S = FALLATTACKRIGHTDOWN Or S = FALLATTACKRIGHT Or S = FALLRIGHT Or FALLMOVERIGHT) Then
                    If (mblnDown = True Or S = FALLATTACKRIGHTDOWN) And S <> FALLATTACKRIGHT Then
                        .pLeft = False
                        ActionFallSwordDown W
                    ElseIf mblnDown <> True Or S = FALLATTACKRIGHT Then
                        .pLeft = False
                        ActionFallSwordOut W
                    End If
                ElseIf (S = ATTACKDOWNRIGHT Or S = ATTACKOUTRIGHT Or S = DUCKRIGHT) Then
                    If (mblnRight = True Or S = ATTACKDOWNRIGHT) And S <> ATTACKOUTRIGHT Then
                        ActionSwordDown W
                    ElseIf mblnRight <> True Or S = ATTACKOUTRIGHT Then
                        ActionSwordOut W
                    End If
                ElseIf (S = ATTACKDOWNLEFT Or S = ATTACKOUTLEFT Or S = DUCKLEFT) Then
                    If (mblnLeft = True Or S = ATTACKDOWNLEFT) And S <> ATTACKOUTLEFT Then
                        ActionSwordDown W
                    ElseIf mblnLeft <> True Or S = ATTACKOUTLEFT Then
                        ActionSwordOut W
                    End If
                ElseIf (S = STANDRIGHT Or S = WALKRIGHT Or S = DASHRIGHT Or S = TURNRIGHT Or S = STOPRIGHT Or S = ATTACKRIGHT Or S = LANDRIGHT Or S = STANDUPRIGHT) Then
                    ActionSword W
                ElseIf (S = STANDLEFT Or S = WALKLEFT Or S = DASHLEFT Or S = TURNLEFT Or S = STOPLEFT Or S = ATTACKLEFT Or S = LANDLEFT Or S = STANDUPLEFT) Then
                    ActionSword W
                End If
            End If
        ' Standing, Stopping, And Idle Animations
        Else
            AlucardGravity S, W
            If .Teeter = True And (S = STOPRIGHT Or S = WALKRIGHT) Then
                ActionTeeter W
            ElseIf .Teeter = True And (S = STOPLEFT Or S = WALKLEFT) Then
                ActionTeeter W
            ElseIf .Teeter = True And (S = STANDRIGHT Or S = DASHRIGHT) Then
                ActionStill W
            ElseIf .Teeter = True And (S = STANDLEFT Or S = DASHLEFT) Then
                ActionStill W
            ElseIf (S = DASHRIGHT Or S = STANDRIGHT) Then
                ActionBob W
            ElseIf (S = DASHLEFT Or S = STANDLEFT) Then
                ActionBob W
            ElseIf S = STOPRIGHT Then
                ActionStop W
            ElseIf S = STOPLEFT Then
                ActionStop W
            ElseIf (S = DUCKRIGHT Or S = STANDUPRIGHT) Then
                ActionStand W
            ElseIf (S = DUCKLEFT Or S = STANDUPLEFT) Then
                ActionStand W
            End If
        End If
        If .DrwWep = False Then .WEP = 0
        'Debug.Print PX(CH) + (-BgX); PY(CH) + (-BgY); BgX; BgY; PX(CH); PY(CH); BgScroll
        'Debug.Print ANI(CH); S; mblnAttack
    End With
End Function

Private Function AlucardGravity(S As Long, W As Long) As Boolean
    With Player
        .Grav = .Grav + 1
        If .Grav >= 200 Then .Grav = 200
        If .FD + 0.2 < 5.4 Then .FD = .FD + 0.2
        If .FS + 0.2 < 8.4 Then .FS = .FS + 0.2
        MoveDown CInt(.FD), CInt(.FS)
        If .WallD = True And .Grav > 20 And (S = FALLRIGHT Or S = LANDRIGHT Or S = FALLMOVERIGHT Or S = JUMPATTACKRIGHT Or S = FALLATTACKRIGHT Or S = FALLATTACKRIGHTDOWN) Then
            .pLeft = False
            ActionHardLand W
            AlucardGravity = True
        ElseIf .WallD = True And .Grav > 20 And (S = FALLLEFT Or S = LANDLEFT Or S = FALLMOVELEFT Or S = JUMPATTACKLEFT Or S = FALLATTACKLEFT Or S = FALLATTACKLEFTDOWN) Then
            .pLeft = True
            ActionHardLand W
            AlucardGravity = True
        ElseIf .WallD = True And (S = FALLRIGHT Or S = LANDRIGHT Or S = FALLMOVERIGHT Or S = JUMPATTACKRIGHT Or S = FALLATTACKRIGHT Or S = FALLATTACKRIGHTDOWN) Then
            .pLeft = False
            ActionLand W
            AlucardGravity = True
        ElseIf .WallD = True And (S = FALLLEFT Or S = LANDLEFT Or S = FALLMOVELEFT Or S = JUMPATTACKLEFT Or S = FALLATTACKLEFT Or S = FALLATTACKLEFTDOWN) Then
            .pLeft = True
            ActionLand W
            AlucardGravity = True
        ElseIf .WallD = False And mblnRight = True And (S = FALLRIGHT Or S = FALLLEFT Or S = FALLMOVERIGHT Or S = FALLMOVELEFT) Then
            .pLeft = False
            ActionAirMove W
        ElseIf .WallD = False And mblnLeft = True And (S = FALLRIGHT Or S = FALLLEFT Or S = FALLMOVERIGHT Or S = FALLMOVELEFT) Then
            .pLeft = True
            ActionAirMove W
        ElseIf .WallD = False And (S = FALLRIGHT Or S = STOPRIGHT Or S = STANDRIGHT Or S = ATTACKRIGHT Or S = ATTACKOUTRIGHT Or S = ATTACKDOWNRIGHT Or S = DUCKRIGHT Or S = DASHRIGHT Or S = WALKRIGHT Or S = TURNRIGHT) Then
            .pLeft = False
            ActionFall W
        ElseIf .WallD = False And (S = FALLLEFT Or S = STOPLEFT Or S = STANDLEFT Or S = ATTACKLEFT Or S = ATTACKOUTLEFT Or S = ATTACKDOWNLEFT Or S = DUCKLEFT Or S = DASHLEFT Or S = WALKLEFT Or S = TURNLEFT) Then
            .pLeft = True
            ActionFall W
        End If
        If .WallD = True Then ' Reset Jumping & Falling Gravity
            .JD = 5.4
            .JS = 8.4
            .FD = 0.4
            .FS = 3.4
        End If
    End With
End Function

Private Function MoveRight(intDistance As Integer, intSpace As Integer) As Integer
    With Player
        BackgroundScroll
        CollisionDetect .WallR, intSpace, 0, COLLISIONRIGHT
        If .WallR = False Or .StairR = True Then
            If ScrlR = True Then
                BgX = BgX - intDistance
                BgX2 = BgX * 2 \ 3 'BgX2 - intDistance + 1
                BgX3 = BgX2 * 2 \ 3
            Else
                .PX = .PX + intDistance
            End If
        End If
        If .StairR = True Then '/up
            If ScrlU = False Then
                .PY = .PY - intDistance
            ElseIf ScrlU = True Then
                BgY = BgY + intDistance
            End If
        ElseIf .StairL = True Then  '\-down
            If ScrlU = False And ScrlD = False Then
                .PY = .PY + intDistance
            ElseIf ScrlD = True Then
                BgY = BgY - intDistance
            End If
        End If
        .StairL = False
        .StairR = False
        BgY2 = BgY * 2 \ 3
        BgY3 = BgY2 * 2 \ 3
    End With
End Function

Private Function MoveLeft(intDistance As Integer, intSpace As Integer) As Integer
    With Player
        BackgroundScroll
        CollisionDetect .WallL, -intSpace, 0, COLLISIONLEFT
        If .WallL = False Or .StairL = True Then
            If ScrlL = True Then
                BgX = BgX + intDistance
                BgX2 = BgX * 2 \ 3 'BgX2 + intDistance - 1
                BgX3 = BgX2 * 2 \ 3
            Else
                .PX = .PX - intDistance
            End If
        End If
        'Add another collision detection if there is a celing at the top of stairs to hit.
        If .StairL = True Then ' \-up
            If ScrlU = False Then
                .PY = .PY - intDistance
            ElseIf ScrlU = True Then
                BgY = BgY + intDistance
            End If
        ElseIf .StairR = True Then '/-down
            If ScrlU = False And ScrlD = False Then
                .PY = .PY + intDistance
            ElseIf ScrlD = True Then
                BgY = BgY - intDistance
            End If
        End If
        .StairL = False
        .StairR = False
        BgY2 = BgY * 2 \ 3
        BgY3 = BgY2 * 2 \ 3
    End With
End Function

Private Function MoveDown(intDistance As Integer, intSpace As Integer) As Integer
    With Player
        BackgroundScroll
        CollisionDetect .WallD, 0, intSpace, COLLISIONDOWN
        If ScrlD = True Then
            If .WallD = False Then
                BgY = BgY - intDistance
            ElseIf .WallD = True Then
                BgY = BgY - .NY
            End If
        ElseIf ScrlD = False Then
            If .WallD = False Then
                .PY = .PY + intDistance
            ElseIf .WallD = True Then
                .PY = .PY + .NY
            End If
        End If
        BgY2 = BgY * 2 \ 3
        BgY3 = BgY2 * 2 \ 3
    End With
End Function

Private Function MoveUp(intDistance As Integer, intSpace As Integer) As Long
    With Player
        BackgroundScroll
        CollisionDetect .WallU, 0, -intSpace, COLLISIONUP
        If ScrlU = True Then
            If .WallU = False Then
                BgY = BgY + intDistance
            ElseIf .WallU = True Then
                .JT = JUMPHEIGHT + 1 ' No More Jumping
                BgY = BgY + .NY
            End If
        ElseIf ScrlU = False Then
            If .WallU = False Then
                .PY = .PY - intDistance
            ElseIf .WallU = True Then
                .JT = JUMPHEIGHT + 1 ' No More Jumping
                .PY = .PY - .NY
            End If
        End If
        BgY2 = BgY * 2 \ 3
        BgY3 = BgY2 * 2 \ 3
    End With
End Function


'/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\
' MOVEMENT
' -------------------------------------------------------------------
' EFFECTS
'\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/


Private Sub MoveEffects()
Dim L As Long, D As String, N As Long, H As String, C As Long
    H = Player.HP ' Make String
    L = Len(H)  ' Didgets of health
    C = 4
    Head(1) = 31 ' Backdrop
    Head(2) = 32 ' Heart
    If HDelay(3) = 2 Then ' Health bar
        HDelay(3) = 0
        Head(3) = Head(3) + 1
        If Head(3) < 34 Or Head(3) > 38 Then Head(3) = 34
    Else
        HDelay(3) = HDelay(3) + 1
        If Head(3) = 0 Then Head(3) = 34
    End If
    Do Until L = 0 ' For each letter
        C = C + 1
        D = Right(H, L)
        N = Left(D, 1) ' The number to evaluate
        Head(C) = N + 21 ' 21 to the first number sprite
        L = L - 1
    Loop
    HeadCount = C
    fX = 0 ' Splash
    If pblnDrawSplash(CH) = True Then Splash
    fX = 1 ' Candles
    Candles
End Sub

Function MoveProjectiles() As Long
'Axe 1 to 12
    With Enemy(CW)
        If .WX + BgX > 320 Or .WX + BgX < 0 Or .WY + BgY < 0 Or .WY + BgY > 240 Then
            If CW = TotlSub Then ' Just Delete
                TotlSub = TotlSub - 1
            Else ' Replace With Last Item, Then Delete Last
                .SWEP = Enemy(TotlSub).SWEP
                .wLeft = Enemy(TotlSub).wLeft
                .WX = Enemy(TotlSub).WX
                .WY = Enemy(TotlSub).WY
                TotlSub = TotlSub - 1
            End If
        Else
            If .SWEP < 1 Or .SWEP > 12 Then
                .SWEP = 1
            End If
            If .wLeft = True Then
                .WX = .WX - 3
            Else
                .WX = .WX + 3
            End If
            If .SWEP >= 0 And .SWEP < 12 Then
                .SWEP = .SWEP + 1
            ElseIf .SWEP = 12 Then
                .SWEP = 1
            End If
        End If
    End With
End Function

Private Function Sword(S As Integer, E As Integer, F As Integer) As Integer
'Start, End, First
    With Player
        If .WEP < S Or .WEP > E Then .WEP = F
        If .WEP >= F And .WEP < E Then
            .WEP = .WEP + 1
        ElseIf .WEP = E Then
            .WEP = .WEP + 1
            .DrwWep = False
        End If
    End With
End Function

Private Sub Splash()
'Splash 11 to 14
    If AniEfct(CH) < 11 Or AniEfct(CH) > 14 Then AniEfct(CH) = 11
    If EfctX(fX) = 0 And EfctY(fX) = 0 Then
        EfctX(fX) = Player.PX - BgX - 10
        EfctY(fX) = Player.PY - BgY - 35
    End If
    If EfctDly(fX) = 3 Then
        EfctDly(fX) = 0
        If AniEfct(fX) >= 11 And AniEfct(fX) < 14 Then
            EfctDone(fX) = False
            AniEfct(fX) = AniEfct(fX) + 1
        ElseIf AniEfct(fX) >= 14 Then
            EfctDone(fX) = True
            EfctX(fX) = 0
            EfctY(fX) = 0
            pblnDrawSplash(fX) = False
            AniEfct(fX) = 0
        End If
    Else
        EfctDly(fX) = EfctDly(fX) + 1
    End If
End Sub

Private Sub Candles()
Dim blnR As Boolean
Static OC(1 To 4) As Long ' Old candles
Static PC As Long ' Currently Used candle sprites
    EfctDly(fX) = EfctDly(fX) + 1
    If EfctDly(fX) >= 10 Then
        EfctDly(fX) = 0
        Randomize
        mintCandle = 6 * Rnd ' Pick A Random Frame Of Animation
        PC = PC + 1
        If PC > 4 Then PC = 1
        OC(PC) = mintCandle
        Do Until mintCandle <> OC(1) And mintCandle <> OC(2) And mintCandle <> OC(3) And mintCandle <> OC(4) And mintCandle <> 0
            mintCandle = CInt(6 * Rnd)
        Loop
        OC(PC) = mintCandle
    End If
End Sub


'/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\
' EFFECTS
' -------------------------------------------------------------------
' ENEMY MOVEMENT
'\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/


Private Sub MoveEnemy()
Dim S As Long
    For CH = 1 To TotlEn
        With Enemy(CH)
            S = .STA
            If S <> WAITING Then EnemyMoveDown DISTANCE, SPACE
            Select Case .EType
                Case 1
                    If .PX + BgX > 420 Or .PX + BgX < -100 Or .PY + BgY < -100 Or .PY + BgY > 340 Then
                        ' It's way off screen, so have it hibernate
                        If S <> WAITING Then EnemyWait .ANI, .STA, .pLeft
                    ElseIf .HURT = True Then
                        EnemyDie .ANI, .STA, .pLeft
                    Else
                        If S = WAITING Or S = UNSPAWNING Then
                            If .PX + BgX < 320 And .PX + BgX > 1 And .PY + BgY < 240 And .PY + BgY > 1 Then
                                CreateZombie .ANI, .STA, .pLeft
                            End If
                        ElseIf S = SPAWNING Then
                            CreateZombie .ANI, .STA, .pLeft
                        ElseIf .pLeft = False And S <> WALKLEFT And S <> DYING And S <> DEAD And S <> UNSPAWNING Then
                            If .WallR = True Or .StairR = True Then
                                .pLeft = True ' Hit wall or stair, turn around
                                .STA = WALKLEFT
                            ElseIf .WallR = False Then
                                Randomize
                                If (.PX + BgX < 320) And (.PX + BgX > 0) And (.PY + BgY > 0) And (.PY + BgY < 240) And (Rnd * 10 > 9.7) Then ' Target Right & In Range
                                    TotlSub = TotlSub + 1
                                    Enemy(TotlSub).wLeft = False
                                    Enemy(TotlSub).WX = .PX
                                    Enemy(TotlSub).WY = .PY
                                Else
                                    EnemyWalk .ANI, .AniDelay, .STA, .pLeft
                                End If
                            End If
                        ElseIf .pLeft = True And S <> WALKRIGHT And S <> DYING And S <> DEAD And S <> UNSPAWNING Then
                            If .WallL = True Or .StairL = True Then
                                .pLeft = False ' Hit wall or stair, turn around
                                .STA = WALKRIGHT
                            ElseIf .WallL = False Then
                                Randomize
                                If (.PX + BgX < 320) And (.PX + BgX > 0) And (.PY + BgY > 0) And (.PY + BgY < 240) And (Rnd * 10 > 9.7) Then ' Target Left & In Range
                                    TotlSub = TotlSub + 1
                                    Enemy(TotlSub).wLeft = True
                                    Enemy(TotlSub).WX = .PX
                                    Enemy(TotlSub).WY = .PY
                                Else
                                    EnemyWalk .ANI, .AniDelay, .STA, .pLeft
                                End If
                            End If
                        End If
                    End If
                Case 2
                    If .PX + BgX > 420 Or .PX + BgX < -100 Or .PY + BgY < -100 Or .PY + BgY > 340 Or .HURT = True Then
                        EnemyDie .ANI, .STA, .pLeft
                    Else
                        'Select Case S
                            'Case iS = DEAD
                            '    StartPosition ANI(CH), STA(CH)
                            'Case iS = SPAWNING
                            '    CreateZombie ANI(CH), STA(CH), pLeft(CH)
                            'Case Is <> WALKLEFT, Is <> DYING, Is <> FALLLEFT, Is <> FALLRIGHT And WallR(CH) = False
                            '    'pLeft(CH) = True
                            '    EnemyWalk ANI(CH), AniDelay(CH), STA(CH), pLeft(CH)
                            'Case Is <> WALKRIGHT, Is <> DYING, Is <> FALLRIGHT, Is <> FALLLEFT And WallL(CH) = False
                                'pLeft(CH) = False
                                BloodyZombieWalk .ANI, .AniDelay, .STA, .pLeft
                        'End Select
                    End If
            End Select
        End With
    Next
End Sub


'/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\
' ENEMY MOVEMENT
' -------------------------------------------------------------------
' TOY BOX
'\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/


Private Function ChangeBackground(intExit As Long) As String
Dim strRecord As String, F As Long
Dim intNull As Integer ' Nothing for LoadSpriteSheet
    'Dump Old Background DCs
    DeleteDC memBgColl
    DeleteDC memBgSpr
    DeleteDC memBgCollision
    DeleteDC memBgPsSpr
    DeleteDC memBgPsMsk
    DeleteDC memBgUpSpr
    DeleteDC memBgUpMsk
    DeleteDC memBgUpColl
    DeleteDC memBgSpr2
    DeleteDC memBgMsk2
    TotlItm = 0
    'Load Images
    'Layer 1 Sprite
    memBgSpr = LoadDC(App.Path & BgPath(intExit) & BgSpr(intExit), True)
    'Layer 2 Sprite and Mask
    If Dir$(App.Path & BgPath(intExit) & BgSpr2(intExit)) <> "" And Len(BgSpr2(intExit)) <> 0 Then
        memBgSpr2 = LoadDC(App.Path & BgPath(intExit) & BgSpr2(intExit), False)
        BgLayer(2) = True
    Else
        BgLayer(2) = False
    End If
    'Layer 3 Sprite and Mask
    If Dir$(App.Path & BgPath(intExit) & BgMsk(intExit)) <> "" And Len(BgMsk(intExit)) <> 0 Then
        memBgMsk2 = LoadDC(App.Path & BgPath(intExit) & BgMsk(intExit), False)
        BgLayer(2) = True
    Else
        BgLayer(2) = False
    End If
    'Layer 4 Sprite and Mask
    If Dir$(App.Path & BgPath(intExit) & BgSpr3(intExit)) <> "" And Len(BgSpr3(intExit)) <> 0 Then
        memBgSpr3 = LoadDC(App.Path & BgPath(intExit) & BgSpr3(intExit), False)
        BgLayer(3) = True
    Else
        BgLayer(3) = False
    End If
    If Dir$(App.Path & BgPath(intExit) & BgMsk2(intExit)) <> "" And Len(BgMsk2(intExit)) <> 0 Then
        memBgMsk3 = LoadDC(App.Path & BgPath(intExit) & BgMsk2(intExit), False)
        BgLayer(3) = True
    Else
        BgLayer(3) = False
    End If
    
    
    memBgColl = LoadDC(App.Path & BgPath(intExit) & BgColl(intExit), False)
    memBgCollision = LoadDC(App.Path & BgPath(intExit) & BgColl(intExit), False)
    If Dir$(App.Path & BgPath(intExit) & BgPsSpr(intExit)) <> "" And Len(BgPsSpr(intExit)) <> 0 Then
        memBgPsSpr = LoadDC(App.Path & BgPath(intExit) & BgPsSpr(intExit), False)
    End If
    If Dir$(App.Path & BgPath(intExit) & BgPsMsk(intExit)) <> "" And Len(BgPsMsk(intExit)) <> 0 Then
        memBgPsMsk = LoadDC(App.Path & BgPath(intExit) & BgPsMsk(intExit), False)
    End If
    If Dir$(App.Path & BgPath(intExit) & BgUpSpr(intExit)) <> "" And Len(BgUpSpr(intExit)) <> 0 Then
        memBgUpSpr = LoadDC(App.Path & BgPath(intExit) & BgUpSpr(intExit), False)
    End If
    If Dir$(App.Path & BgPath(intExit) & BgUpMsk(intExit)) <> "" And Len(BgUpMsk(intExit)) <> 0 Then
        memBgUpMsk = LoadDC(App.Path & BgPath(intExit) & BgUpMsk(intExit), False)
    End If
    If Dir$(App.Path & BgPath(intExit) & BgUpColl(intExit)) <> "" And Len(BgUpColl(intExit)) <> 0 Then
        memBgUpColl = LoadDC(App.Path & BgPath(intExit) & BgUpColl(intExit), False)
    End If
    'Update Scrolling Information
    BgCurrent = BgSpr(intExit)
    BgScroll = BgType(intExit)
    BgExitTxt = BgExitNxt(intExit)
    BgX = BgNewX(intExit)
    BgY = BgNewY(intExit)
    BgX2 = BgX * 2 \ 3
    BgY2 = BgY * 2 \ 3
    BgX3 = BgX2 * 2 \ 3
    BgY3 = BgY2 * 2 \ 3
    Player.PX = BgStartX(intExit)
    Player.PY = BgStartY(intExit)
    'Load Passable Objects
    If Dir$(App.Path & BgPath(intExit) & BgPsTxt(intExit)) <> "" And Len(BgPsTxt(intExit)) <> 0 Then
        Open App.Path & BgPath(intExit) & BgPsTxt(intExit) For Input As #10
            Do Until EOF(10)
                F = F + 1
                PsDth(F) = False
                LoadSpriteSheet PsXpos(F), PsYpos(F), PsXp(F), PsYp(F), PsWp(F), PsHp(F), PItm(F), PItmFat(F), True
            Loop
        Close #10
    Else
        For F = 1 To TotlCand
            PsXpos(F) = 0: PsYpos(F) = 0: PsXp(F) = 0: PsYp(F) = 0: PsWp(F) = 0: PsHp(F) = 0
        Next
        TotlCand = 0
    End If
    'Load UnPassable Objects
    F = 0
    If Dir$(App.Path & BgPath(intExit) & BgUpTxt(intExit)) <> "" And Len(BgUpTxt(intExit)) <> 0 Then
        Open App.Path & BgPath(intExit) & BgUpTxt(intExit) For Input As #10
            Do Until EOF(10)
                F = F + 1
                LoadSpriteSheet UpXpos(F), UpYpos(F), UpXp(F), UpYp(F), UpWp(F), UpHp(F), intNull, intNull, False
            Loop
        Close #10
    Else
        For F = 1 To 10
            UpXpos(F) = 0: UpYpos(F) = 0: UpXp(F) = 0: UpYp(F) = 0: UpWp(F) = 0: UpHp(F) = 0
        Next
        TotlUp = 0
    End If
    'Destroy characters
    For F = 1 To TotlEn
        Enemy(F).STA = 0
    Next
    'Load Characters
    ChangeEnemies intExit
    'Whipe out old background data
    For F = 1 To BgExitTotal
        BgPath(F) = 0
        BgColl(F) = ""
        BgCurExit(F) = 0
        BgExitNxt(F) = 0
        BgSpr(F) = ""
        BgMsk(F) = ""
        BgSpr2(F) = ""
        BgMsk2(F) = ""
        BgSpr3(F) = ""
        BgMsk3(F) = ""
        BgPsSpr(F) = ""
        BgPsMsk(F) = ""
        BgUpSpr(F) = ""
        BgUpMsk(F) = ""
        BgUpColl(F) = ""
        BgPsTxt(F) = ""
        BgUpTxt(F) = ""
        BgExitX(F) = 0
        BgExitY(F) = 0
        BgStartX(F) = 0
        BgStartY(F) = 0
        BgNewX(F) = 0
        BgNewY(F) = 0
        BgType(F) = 0
        BgExitType(F) = 0
    Next
End Function

Private Sub FinalCalculations()
Dim W As Long, H As Long ' Hazards
Dim C As Boolean ' Collision
    MoveEffects
    H = 0
    With Player
        W = .ANI
        ' Create Alucard Left Xpos
        .LXP = .CEN - (.Widthp(W) + .Xpos(W) - .CEN)
        ' Create a weapon's left Xpos
        .SXP = 70 - .WepW(.WEP) - .WepXpos(.WEP)
    End With
    ' Create Enemy Left Xpos
    For H = 1 To TotlEn
        With Enemy(H)
            W = .ANI
            .LXP = .CEN - (EWidthp(W, .EType) + EXpos(W, .EType) - .CEN)
        End With
    Next
    'Create Projectile Weapons Left Xpos
    For H = 1 To TotlSub
        With Enemy(H)
            W = .SWEP
            .WCEN = 20 ' Hard code these values for each type of weapon
            .WXP = .WCEN - (EWepW(W) + EWepXpos(W) - .WCEN)
        End With
    Next
    ' Move Projectile Weapons & Collisions
    For CW = 1 To TotlSub
        C = False
        MoveProjectiles
        WCollisionDetect C, CW
        If C = True Then Player.HURT = True
    Next
    ' See if weapon hit any projectiles
    'For CW = 1 To TotlSub
    '    C = False
    '    If Player.DrwWep = True Then DCollisionDetect C, CW
    '    If C = True Then
    '        Player.wLeft(CW) = Player.pLeft
    '    End If
    'Next
    ' See if Alucard touched any enemies
    For H = 1 To TotlEn
        If Player.HURT = True Then Exit For
        SCollisionDetect Player.HURT, H, Enemy(H).EType
    Next
    ' See if weapon hit any objects
    For H = 1 To TotlCand
        If Player.DrwWep = True Then OCollisionDetect H
    Next
    ' See if Alucard touched any Items
    For H = 1 To TotlItm
        C = False
        CCollisionDetect C, H
        If C = False Then
            'Item Gravity
            C = False
            ICollisionDetect C, H ' Check for Floor-Item Collision
            If C = False Then
                IYpos(H) = IYpos(H) + ItmFat(H)
            End If
        End If
    Next
End Sub

Private Sub ResetTurn()
    With Player
        If .STA = TURNRIGHT Then
            .STA = STANDRIGHT
        ElseIf .STA = TURNLEFT Then
            .STA = STANDLEFT
        End If
    End With
End Sub

Private Sub CheckBackground()
Dim E As Long, B As Boolean ' Loop & Change Background
Dim M As Long, N As Long ' Music Status & New Song
    'Check for Change
    For E = 1 To BgExitTotal
        With Player
            Select Case BgExitType(E)
                Case 0 'Right x > ,y >
                    If .PX + (-BgX) >= BgExitX(E) And .PY + (-BgY) >= BgExitY(E) Then: B = True: Exit For
                Case 1 ' Left x < ,y >
                    If .PX + (-BgX) <= BgExitX(E) And .PY + (-BgY) >= BgExitY(E) Then: B = True: Exit For
                Case 2 ' Top x > ,y <
                    If .PX + (-BgX) >= BgExitX(E) And .PY + (-BgY) <= BgExitY(E) Then: B = True: Exit For
                Case 3 ' Down x > ,y >
                    If .PX + (-BgX) >= BgExitX(E) And .PY + (-BgY) >= BgExitY(E) Then: B = True: Exit For
            End Select
        End With
    Next
    'Change Background
    If B = True Then
        If blnMsic = True Then M = BgMsic.GetStatus
        N = LoadSound(E, "", M)
        ChangeBackground E
        UpdateBackgroundCollision
        LoadNextBackground
        If M = 0 Or N = 1 Then ' No Music Playing & No New Music
            If blnMsic = True Then BgMsic.Play DSBPLAY_LOOPING
        End If
    End If
End Sub

Private Sub BackgroundScroll()
Dim BgSrlH As Long
Dim BgSrlV As Long
Const SCROLLHORIZONTAL = 0
Const SCROLLVERTICAL = 1
Const SCROLLALL = 2
'Const SCROLLNONE = 3
    With Player
        BgSrlH = (BgWidth - SCREENWIDTH - 7)
        BgSrlV = (BgHeight - SCREENHEIGHT - 7)
        ScrlL = False
        ScrlR = False
        ScrlU = False
        ScrlD = False
        'Horizontal
        If BgScroll = SCROLLHORIZONTAL Or BgScroll = SCROLLALL Then
            If -BgX > BgSrlH Then 'Room to scroll left
                If .PX < 135 And .PX > 125 Then 'Near Center Screen
                    ScrlL = True
                End If
            ElseIf -BgX < 6 Then 'Room to scroll right
                If .PX < 135 And .PX > 125 Then
                    ScrlR = True
                End If
            ElseIf .PX < 135 And .PX > 125 Then
                ScrlL = True
                ScrlR = True
            End If
        End If
        'Vertical
        If BgScroll = SCROLLVERTICAL Or BgScroll = SCROLLALL Then
            If -BgY > BgSrlV Then 'Room to scroll down
                If .PY < 105 And .PY > 95 Then 'Near Center Screen
                    ScrlU = True
                End If
            ElseIf -BgY < 6 Then ' Room to scroll up
                If .PY < 105 And .PY > 95 Then
                    ScrlD = True
                End If
            ElseIf .PY < 105 And .PY > 95 Then
                ScrlU = True
                ScrlD = True
            End If
        End If
    End With
End Sub

Private Sub UpdateBackgroundCollision()
Dim Obj As Long
    For Obj = 1 To TotlUp
        BitBlt memBgCollision, UpXpos(Obj), UpYpos(Obj), UpWp(Obj), UpHp(Obj), memBgUpColl, UpXp(Obj), UpYp(Obj), &HCC0020 ' SRCCOPY
    Next
End Sub

Private Sub GameOverScreen()
Dim L As Long
    Do Until mblnEndProgram = True Or GameOver = False
        DrawScreen
        DoEvents
        With Me
            .Font.Name = "Verdana"
            'Draw Game Over
            For L = 1 To 4
                .ForeColor = 65280
                .FontSize = 40
                Select Case L
                    Case 1: .CurrentX = 4: .CurrentY = 50: Me.Print "Game Over"
                    Case 2: .CurrentX = 6: .CurrentY = 50: Me.Print "Game Over"
                    Case 3: .CurrentX = 5: .CurrentY = 51: Me.Print "Game Over"
                    Case 4: .CurrentX = 5: .CurrentY = 49: Me.Print "Game Over"
                End Select
            Next
            'Draw Restart
            .FontSize = 20
            If mblnLeft = True Then
                For L = 1 To 4
                    Select Case L
                        Case 1: .CurrentX = 51: .CurrentY = 120: Me.Print "Restart"
                        Case 2: .CurrentX = 49: .CurrentY = 120: Me.Print "Restart"
                        Case 3: .CurrentX = 50: .CurrentY = 121: Me.Print "Restart"
                        Case 4: .CurrentX = 50: .CurrentY = 119: Me.Print "Restart"
                    End Select
                Next
            'Draw Quit
            ElseIf mblnRight = True Then
                For L = 1 To 4
                    Select Case L
                        Case 1: .CurrentX = 184: .CurrentY = 120: Me.Print "Quit"
                        Case 2: .CurrentX = 186: .CurrentY = 120: Me.Print "Quit"
                        Case 3: .CurrentX = 185: .CurrentY = 119: Me.Print "Quit"
                        Case 4: .CurrentX = 185: .CurrentY = 121: Me.Print "Quit"
                    End Select
                Next
            'Draw Space
            Else
                .FontSize = 30
                For L = 1 To 4
                    Select Case L
                        Case 1: .CurrentX = 159: .CurrentY = 100: Me.Print "|"
                        Case 2: .CurrentX = 161: .CurrentY = 100: Me.Print "|"
                        Case 3: .CurrentX = 160: .CurrentY = 99: Me.Print "|"
                        Case 4: .CurrentX = 160: .CurrentY = 101: Me.Print "|"
                    End Select
                Next
            End If
            .ForeColor = 0
            .CurrentX = 5: .CurrentY = 50: .FontSize = 40: Me.Print "Game Over"
            .CurrentX = 160: .CurrentY = 100: .FontSize = 30: Me.Print "|"
            .FontSize = 20
            .CurrentX = 50: .CurrentY = 120: Me.Print "Restart"
            .CurrentX = 185: .CurrentY = 120: Me.Print "Quit"
        End With
        'Exit loop to restart game
        If (mblnLeft = True Or mblnRight = True) And (mblnAttack = True Or mblnUp = True) Then GameOver = False
    Loop
    If mblnLeft = True And (mblnAttack = True Or mblnUp = True) Then ' Restart Game
        DumpDCs
        TotlItm = 0
        BgX = 0: BgY = 0
        CH = 0 ' Alucard
        Player.STA = STANDRIGHT
        Player.HURT = False
        LoadUpSprites
        UpdateBackgroundCollision
        LoadNextBackground
        GameLoop
    ElseIf (mblnRight = True And (mblnAttack = True Or mblnUp = True)) Or mblnEndProgram = True Then ' Exit Game
        mblnEndProgram = True
        DumpDCs
        DumpRequiredDCs
        frmChangeRez.EndRez
        Unload frmChangeRez
        Unload frmCollision
        Unload Me
    End If
End Sub


'/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\
' TOY BOX
' -------------------------------------------------------------------
' LOADING
'\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/-\/


Private Function ChangeEnemies(E As Long) As Long
Dim strRecord As String, S As String, M As String ' Sprite & Mask
Dim T As Long ' Type of sprite sheet
Dim intNull As Integer ' Nothing for LoadSpriteSheet
Dim intAnim As Integer, intType As Integer
    Open App.Path & BgPath(E) & BgChTxt(E) For Input As #20
        Input #20, TotlEn
        Do Until EOF(20)
            Input #20, strRecord
            'Load Enemy Sprites & Masks
            If strRecord = "Sprites" Then
                Do Until S = "End" Or M = "End"
                    Input #20, S
                    Input #20, M
                    Input #20, T
                    If S <> "End" And M <> "End" Then
                        memESpr(T) = LoadDC(App.Path & S, False)
                        memEMsk(T) = LoadDC(App.Path & M, False)
                    End If
                Loop
            'Load Coordinates
            ElseIf strRecord = "Coordinates" Then
                Do Until strRecord = "End"
                    Input #20, strRecord
                    If strRecord <> "End" Then
                        intAnim = 0
                        intType = intType + 1
                        Open App.Path & strRecord For Input As #10
                            Do Until EOF(10)
                                intAnim = intAnim + 1
                                LoadSpriteSheet EXpos(intAnim, intType), EYpos(intAnim, intType), EXp(intAnim, intType), EYp(intAnim, intType), EWidthp(intAnim, intType), EHeightp(intAnim, intType), intNull, intNull, False
                            Loop
                        Close #10
                    End If
                Loop
            'Load Enemy Stats
            Else
                CH = strRecord
                'Alucard is included and he isn't spawning, so exit Do Loop.
                If CH = 0 And Player.PX <> 0 And Player.PY <> 0 Then Exit Do
                Input #20, Enemy(CH).BL
                Input #20, Enemy(CH).BT
                Input #20, Enemy(CH).BW
                Input #20, Enemy(CH).BH
                Input #20, Enemy(CH).CEN
                Input #20, Enemy(CH).PX
                Input #20, Enemy(CH).PY
                Input #20, Enemy(CH).HP
                Input #20, Enemy(CH).AP
                Input #20, Enemy(CH).DP
                Input #20, Enemy(CH).EType
            End If
            'Set Enemy as New
            For CH = 1 To TotlEn
                Enemy(CH).STA = WAITING
                Enemy(CH).ANI = 0
                Enemy(CH).HURT = False
            Next
        Loop
    Close #20
End Function

Private Sub LoadRequiredSprites()
Dim F As Long
Dim strRecord As String
Dim intFile As Integer
Dim Pixel(34000) As Long
Dim PS As Long
Dim PE As Long
Dim intNull As Integer ' Nothing for LoadSpriteSheet
    memHeadSpr = LoadDC(App.Path & "\Sprites\headsupSpr.gif", False)
    memHeadMsk = LoadDC(App.Path & "\Sprites\headsupMsk.gif", False)
    Player.memSpr = LoadDC(App.Path & "\Sprites\alucardSpr.gif", False)
    Player.memMsk = LoadDC(App.Path & "\Sprites\alucardMsk.gif", False)
    Player.memHurt = LoadDC(App.Path & "\Sprites\alucardHrt.gif", False)
    memItmSpr = LoadDC(App.Path & "\Sprites\itemsSpr.gif", False)
    memItmMsk = LoadDC(App.Path & "\Sprites\itemsMsk.gif", False)
    memSubSpr = LoadDC(App.Path & "\Sprites\subSpr.gif", False)
    memSubMsk = LoadDC(App.Path & "\Sprites\subMsk.gif", False)
    'Required Coordinate files
    intFile = 0
    Open App.Path & "\Sprites\headsup.txt" For Input As #10
        Do Until EOF(10)
            intFile = intFile + 1
            LoadSpriteSheet HXpos(intFile), HYpos(intFile), HXp(intFile), HYp(intFile), HWidth(intFile), HHeight(intFile), intNull, intNull, False
        Loop
    Close #10
    intFile = 0
    Open App.Path & "\Sprites\itemsSpr.txt" For Input As #10
        Do Until EOF(10)
            intFile = intFile + 1
            LoadSpriteSheet IXpos(intFile), IYpos(intFile), IXp(intFile), IYp(intFile), IWp(intFile), IHp(intFile), intNull, intNull, False
        Loop
    Close #10
    intFile = 0
    Open App.Path & "\Sprites\alucardSpr.txt" For Input As #10
        With Player
            Do Until EOF(10)
                intFile = intFile + 1
                LoadSpriteSheet .Xpos(intFile), .Ypos(intFile), .Xp(intFile), .Yp(intFile), .Widthp(intFile), .Heightp(intFile), intNull, intNull, False
            Loop
        End With
    Close #10
    intFile = 0
    Open App.Path & "\Sprites\subSpr.txt" For Input As #10
        Do Until EOF(10)
            intFile = intFile + 1
            LoadSpriteSheet EWepXpos(intFile), EWepYpos(intFile), EWepX(intFile), EWepY(intFile), EWepW(intFile), EWepH(intFile), intNull, intNull, False
        Loop
    Close #10
    CH = 0
    Open App.Path & "\Sprites\alucardCol.txt" For Input As #10
        Do Until EOF(10)
            Input #10, strRecord
            If strRecord = "F" Then
                Input #10, F
                Input #10, CollS(F)
                Input #10, CollE(F)
                Input #10, strRecord
            End If
            PS = CollS(F)
            PE = CollE(F)
            If Pixel(F) >= PS And Pixel(F) < PE Then
                Pixel(F) = Pixel(F) + 1
            Else
                Pixel(F) = PS
            End If
            CollX(Pixel(F)) = strRecord
            Input #10, strRecord
            CollY(Pixel(F)) = strRecord
        Loop
    Close #10
End Sub

Private Function LoadDC(strImg As String, setBG As Boolean) As Long
Dim PreBMP As Long, objBMP As Long, BMP As BITMAP, screenDC As Long
    screenDC = GetDC(memScreen)
    LoadDC = CreateCompatibleDC(screenDC)
    'Clean And Draw DC
    If Len(strImg) <> 0 Then
        PreBMP = SelectObject(LoadDC, LoadPicture(App.Path & "\Backgrounds\blank.bmp"))
        PreBMP = SelectObject(LoadDC, LoadPicture(strImg))
    Else
        PreBMP = SelectObject(LoadDC, LoadPicture(App.Path & "\Backgrounds\blank.bmp"))
    End If
    'Get Height And Width
    If setBG = True Then
        objBMP = GetObject(LoadPicture(strImg), Len(BMP), BMP)
        BgWidth = BMP.bmWidth
        BgHeight = BMP.bmHeight
    End If
    'Delete Temporary Objects
    DeleteDC screenDC
    DeleteObject objBMP
    DeleteObject PreBMP
End Function

Private Sub LoadNextBackground()
Dim E As Integer
Dim intExit As Integer
Dim strCurrent As String
    Open App.Path & BgPath(0) & BgExitTxt For Input As #10
        Input #10, BgExitTotal
        Do Until EOF(10)
            E = E + 1
            Input #10, BgPath(E)
            Input #10, BgCurrent
            Input #10, BgColl(E)
            Input #10, BgCurExit(E)
            Input #10, BgExitNxt(E)
            Input #10, BgSpr(E)
            Input #10, BgMsk(E)
            Input #10, BgSpr2(E)
            Input #10, BgMsk2(E) ''''
            Input #10, BgSpr3(E) ''''
            Input #10, BgMsk3(E)
            Input #10, BgPsSpr(E)
            Input #10, BgPsMsk(E)
            Input #10, BgUpSpr(E)
            Input #10, BgUpMsk(E)
            Input #10, BgUpColl(E)
            Input #10, BgPsTxt(E)
            Input #10, BgUpTxt(E)
            Input #10, BgChTxt(E)
            Input #10, BgMuTxt(E)
            Input #10, BgExitX(E)
            Input #10, BgExitY(E)
            Input #10, BgStartX(E)
            Input #10, BgStartY(E)
            Input #10, BgNewX(E)
            Input #10, BgNewY(E)
            Input #10, BgType(E)
            Input #10, BgExitType(E)
        Loop
    Close #10
End Sub

Private Sub LoadUpSprites()
Dim F As Long
Dim strRecord As String
Dim intFile As Integer
Dim Pixel(34000) As Long
Dim PS As Long
Dim PE As Long
Dim intNull As Integer ' Nothing for LoadSpriteSheet
    Open App.Path & "\save.txt" For Input As #20
        Input #20, strRecord
        BgScroll = strRecord
        Input #20, strRecord
        For F = (strRecord + 1) To 4 ' Turn Off Unused Background Layers
            BgLayer(F) = False
        Next
        Input #20, strRecord
        memBgSpr = LoadDC(App.Path & strRecord, True)
        Input #20, strRecord
        memBgColl = LoadDC(App.Path & strRecord, False)
        memBgCollision = LoadDC(App.Path & strRecord, False)
        Input #20, strRecord ' passSpr
        If Len(strRecord) <> 0 Then memBgPsSpr = LoadDC(App.Path & strRecord, False)
        Input #20, strRecord ' passMsk
        If Len(strRecord) <> 0 Then memBgPsMsk = LoadDC(App.Path & strRecord, False)
        Input #20, strRecord
        memBgUpSpr = LoadDC(App.Path & strRecord, False)
        Input #20, strRecord
        memBgUpMsk = LoadDC(App.Path & strRecord, False)
        Input #20, strRecord
        memBgUpColl = LoadDC(App.Path & strRecord, False)
        Input #20, strRecord
        LoadSound 0, strRecord, 0 ' Loads Music Fresh
        Input #20, strRecord
        Player.memWepSpr = LoadDC(App.Path & strRecord, False)
        Input #20, strRecord
        Player.memWepMsk = LoadDC(App.Path & strRecord, False)
        'Background Coordinates Files
        Input #20, strRecord
        If Len(strRecord) <> 0 Then ' passTxt
            Open App.Path & strRecord For Input As #10
                Do Until EOF(10)
                    intFile = intFile + 1
                    PsDth(intFile) = False
                    LoadSpriteSheet PsXpos(intFile), PsYpos(intFile), PsXp(intFile), PsYp(intFile), PsWp(intFile), PsHp(intFile), intNull, intNull, False
                Loop
            Close #10
        End If
        intFile = 0
        Input #20, strRecord
        Open App.Path & strRecord For Input As #10
            Do Until EOF(10)
                intFile = intFile + 1
                LoadSpriteSheet UpXpos(intFile), UpYpos(intFile), UpXp(intFile), UpYp(intFile), UpWp(intFile), UpHp(intFile), intNull, intNull, False
            Loop
        Close #10
        intFile = 0
        Input #20, strRecord ' File Path
        BgPath(intFile) = strRecord
        Input #20, strRecord ' File containing exit information
        BgExitTxt = strRecord
        intFile = 0
        Input #20, strRecord
        'Sword Coordinates
        Open App.Path & strRecord For Input As #10
            With Player
                Do Until EOF(10)
                    intFile = intFile + 1
                    LoadSpriteSheet .WepXpos(intFile), .WepYpos(intFile), .WepX(intFile), .WepY(intFile), .WepW(intFile), .WepH(intFile), intNull, intNull, False
                Loop
            End With
        Close #10
        'Sword Collision
        Input #20, strRecord
        Open App.Path & strRecord For Input As #10
            Do Until EOF(10)
                Input #10, strRecord
                If strRecord = "F" Then
                    Input #10, F
                    Input #10, WepCollS(F)
                    Input #10, WepCollE(F)
                    Input #10, strRecord
                End If
                PS = WepCollS(F)
                PE = WepCollE(F)
                If Pixel(F) >= PS And Pixel(F) < PE Then
                    Pixel(F) = Pixel(F) + 1
                Else
                    Pixel(F) = PS
                End If
                WepCollX(Pixel(F)) = strRecord
                Input #10, strRecord
                WepCollY(Pixel(F)) = strRecord
            Loop
        Close #10
        'Character Background Collision Boxes
        'Don't include Alucard unless in a save room. Otherwise
        'this will fuck up his position when entering a new room.
        intFile = 0
        Input #20, strRecord
        Open App.Path & strRecord For Input As #10
            Input #10, TotlEn
            Do Until EOF(10)
                Input #10, CH
                If CH = 0 Then
                    With Player
                        Input #10, .BL
                        Input #10, .BT
                        Input #10, .BW
                        Input #10, .BH
                        Input #10, .CEN
                        Input #10, .PX
                        Input #10, .PY
                        Input #10, .HP
                        Input #10, .AP
                        Input #10, .DP
                        Input #10, strRecord
                    End With
                Else
                    With Enemy(CH)
                        Input #10, .BL
                        Input #10, .BT
                        Input #10, .BW
                        Input #10, .BH
                        Input #10, .CEN
                        Input #10, .PX
                        Input #10, .PY
                        Input #10, .HP
                        Input #10, .AP
                        Input #10, .DP
                        Input #10, strRecord
                    End With
                End If
            Loop
            If CH = 0 Then Player.MaxHP = Player.HP
        Close #10
    Close #20
    If blnMsic = True Then BgMsic.Play DSBPLAY_LOOPING ' Play Music
End Sub

Private Function LoadSpriteSheet(gXpos As Integer, gYpos As Integer, gX As Integer, gY As Integer, gW As Integer, gH As Integer, gI As Integer, gF As Integer, blnItm As Boolean) As Integer
Dim intRowNumber As Integer
Dim strR As String
Dim strFirst As String
    Input #10, strFirst ' Type or # 1
    If strFirst = "candles" Then
        Input #10, strR ' # Count
        TotlCand = strR
        Input #10, strR ' # 1
    ElseIf strFirst = "unpassable" Then
        Input #10, strR ' # Count
        TotlUp = strR
        Input #10, strR ' # 1
    Else
        strR = strFirst
    End If
    intRowNumber = strR
    Input #10, strR ' # 2
    gXpos = strR
    Input #10, strR ' # 3
    gYpos = strR
    Input #10, strR ' # 4
    gX = strR
    Input #10, strR ' # 5
    gY = strR
    Input #10, strR ' # 6
    gW = strR
    Input #10, strR ' # 7
    gH = strR
    If blnItm = True Then ' Destruction Item
        Input #10, strR
        gI = strR
        Input #10, strR
        gF = strR
    End If
End Function

Private Sub DumpRequiredDCs()
Dim F As Long
    ' Whipe out DCs
    DeleteDC memBlank
    DeleteDC memBuffer
    DeleteDC memCollEnv
    DeleteDC memCollCh
    DeleteDC memHeadSpr
    DeleteDC memHeadMsk
    DeleteDC Player.memSpr
    DeleteDC Player.memMsk
    DeleteDC Player.memHurt
    DeleteDC memItmSpr
    DeleteDC memItmMsk
    ' Whipe out Sound
    For F = 1 To 20
        Set SndEft(F) = Nothing
        Set ESndEft(F) = Nothing
    Next
    Set BgMsic = Nothing
    Set DX = Nothing
    Set DS = Nothing
End Sub

Private Sub DumpDCs()
Dim F As Long
    ' Whipe out DCs
    DeleteDC memBgColl
    DeleteDC Player.memWepSpr
    DeleteDC Player.memWepMsk
    For F = 1 To 100
        DeleteDC memESpr(F)
        DeleteDC memEMsk(F)
    Next
    DeleteDC memBgSpr
    DeleteDC memBgCollision
    DeleteDC memBgPsSpr
    DeleteDC memBgPsMsk
    DeleteDC memBgUpSpr
    DeleteDC memBgUpMsk
    DeleteDC memBgUpColl
    DeleteDC memBgSpr2
    DeleteDC memBgMsk2
End Sub

Private Function LoadSound(N As Long, MU As String, Status As Long) As Long
Dim F As String ' Files
Dim S As Long, E As Long ' Start & End
    If N <> 0 Then
        Open App.Path & BgPath(0) & BgMuTxt(N) For Input As #10 ' From Exit Number
    Else
        Open App.Path & MU For Input As #10 ' From File Name
    End If
        Do Until EOF(10)
            Input #10, F
            If F = "Alucard" Then
                Input #10, E
                For S = 1 To E
                    Input #10, F
                    If Dir$(App.Path & F) <> "" Then
                        blnSndFX(S) = True
                        Set SndEft(S) = Nothing
                        Set SndEft(S) = DS.CreateSoundBufferFromFile(App.Path & F, SndDsc)
                    Else
                        blnSndFX(S) = False
                    End If
                Next
            ElseIf F = "Enemy" Then
                Input #10, E
                For S = 1 To E
                    Input #10, F
                    If Dir$(App.Path & F) <> "" Then
                        EblnSndFX(S) = True
                        Set ESndEft(S) = Nothing
                        Set ESndEft(S) = DS.CreateSoundBufferFromFile(App.Path & F, SndDsc)
                    Else
                        EblnSndFX(S) = False
                    End If
                Next
            ElseIf F = "Music" Then
                Input #10, F
                If Status = 0 Then ' Load Music If There Is None Playing
                    If Dir$(App.Path & F) <> "" Then
                        blnMsic = True
                        Set BgMsic = Nothing
                        Set BgMsic = DS.CreateSoundBufferFromFile(App.Path & F, SndDsc)
                    Else
                        blnMsic = False
                    End If
                End If
            ElseIf F = "MusicNew" Then ' Load Music Regardless If Music Is Playing
                Input #10, F
                If Dir$(App.Path & F) <> "" Then
                    blnMsic = True
                    Set BgMsic = Nothing
                    Set BgMsic = DS.CreateSoundBufferFromFile(App.Path & F, SndDsc)
                    LoadSound = 1 ' Return So Next Procedure Can Force Play
                Else
                    blnMsic = False
                End If
            End If
        Loop
    Close #10
End Function


'/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\-/\
' LOADING
' -------------------------------------------------------------------
