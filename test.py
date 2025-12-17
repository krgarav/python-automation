def build_space_ppt(template_ppt_path, space_data, prs, selected_style, event):

    insert_index = 4
    last_floorplan_layout = None   # track only floorplans

    for space, files in space_data.items():
        print(f"\n📂 Building slides for space: {space}")

        pictures = files.get("pictures", [])
        floorplans = files.get("floor_plans", [])
        elevations = files.get("elevations", [])
        inspiration = files.get("inspiration", [])

        # ---------------------------------------------------------
        # 1️⃣ FLOORPLAN SLIDES (floorplans + pictures)
        # ---------------------------------------------------------
        if pictures and floorplans and elevations:

            # remember the layout for fallback use later
            last_floorplan_layout = floorplans[0]

            # Floorplan/Elevation combo builder (your custom module)
            insert_index = generate_floorplan_elevation_slides(
                "templates/floorplan.pptx",
                floorplans,
                pictures,
                prs,
                insert_index,
                event,
                space      # passed correctly
            )

            # Content layout slides using the actual template
            insert_index = generate_layout_content_slides(
                "templates/imageslide.pptx",
                floorplans,
                pictures,
                prs,
                insert_index,
                event
            )
            
            insert_index = generate_floorplan_elevation_slides(
                "templates/Elevation.pptx",
                elevations,
                pictures,
                prs,
                insert_index,
                event,
                space_name=space
            )
            
            template_ppt_inspiration = "templates/inspiration_slides_template1.pptx"

            insert_index = generate_inspiration_slides(
                template_ppt_inspiration,
                inspiration,
                prs,
                insert_index,
                event
            )
            


        if pictures and not floorplans and elevations:


            if last_floorplan_layout:
                print(f"✔ Using previous FLOORPLAN layout: {last_floorplan_layout}")
                insert_index = generate_layout_content_slides(
                    "templates/floorplan.pptx",
                    [last_floorplan_layout],
                    pictures,
                    prs,
                    insert_index,
                    event
                )
                insert_index = generate_layout_content_slides(
                    "templates/imageslide.pptx",
                    [last_floorplan_layout],
                    pictures,
                    prs,
                    insert_index,
                    event
                )
            
            insert_index = generate_floorplan_elevation_slides(
                "templates/Elevation.pptx",
                elevations,
                pictures,
                prs,
                insert_index,
                event,
                space_name=space
            )
            
            template_ppt_inspiration = "templates/inspiration_slides_template1.pptx"

            insert_index = generate_inspiration_slides(
                template_ppt_inspiration,
                inspiration,
                prs,
                insert_index,
                event
            )
            
         if pictures and not floorplans and elevations:


            if last_floorplan_layout:
                print(f"✔ Using previous FLOORPLAN layout: {last_floorplan_layout}")
                insert_index = generate_layout_content_slides(
                    "templates/floorplan.pptx",
                    [last_floorplan_layout],
                    pictures,
                    prs,
                    insert_index,
                    event
                )
                insert_index = generate_layout_content_slides(
                    "templates/imageslide.pptx",
                    [last_floorplan_layout],
                    pictures,
                    prs,
                    insert_index,
                    event
                )
            
            insert_index = generate_floorplan_elevation_slides(
                "templates/Elevation.pptx",
                elevations,
                pictures,
                prs,
                insert_index,
                event,
                space_name=space
            )
            
            template_ppt_inspiration = "templates/inspiration_slides_template1.pptx"

            insert_index = generate_inspiration_slides(
                template_ppt_inspiration,
                inspiration,
                prs,
                insert_index,
                event
            )
            
            
          
          
          
          
          
          
          
         if pictures and not floorplans and elevations:


            if last_floorplan_layout:
                print(f"✔ Using previous FLOORPLAN layout: {last_floorplan_layout}")
                insert_index = generate_layout_content_slides(
                    "templates/floorplan.pptx",
                    [last_floorplan_layout],
                    pictures,
                    prs,
                    insert_index,
                    event
                )
                insert_index = generate_layout_content_slides(
                    "templates/imageslide.pptx",
                    [last_floorplan_layout],
                    pictures,
                    prs,
                    insert_index,
                    event
                )
            
            insert_index = generate_floorplan_elevation_slides(
                "templates/Elevation.pptx",
                elevations,
                pictures,
                prs,
                insert_index,
                event,
                space_name=space
            )
            
            template_ppt_inspiration = "templates/inspiration_slides_template1.pptx"

            insert_index = generate_inspiration_slides(
                template_ppt_inspiration,
                inspiration,
                prs,
                insert_index,
                event
            )
            
         if pictures and not floorplans and elevations:


            if last_floorplan_layout:
                print(f"✔ Using previous FLOORPLAN layout: {last_floorplan_layout}")
                insert_index = generate_layout_content_slides(
                    "templates/floorplan.pptx",
                    [last_floorplan_layout],
                    pictures,
                    prs,
                    insert_index,
                    event
                )
                insert_index = generate_layout_content_slides(
                    "templates/imageslide.pptx",
                    [last_floorplan_layout],
                    pictures,
                    prs,
                    insert_index,
                    event
                )
            
            insert_index = generate_floorplan_elevation_slides(
                "templates/Elevation.pptx",
                elevations,
                pictures,
                prs,
                insert_index,
                event,
                space_name=space
            )
            
            template_ppt_inspiration = "templates/inspiration_slides_template1.pptx"

            insert_index = generate_inspiration_slides(
                template_ppt_inspiration,
                inspiration,
                prs,
                insert_index,
                event
            )
            
           
          
          if pictures and floorplans and elevations:

            # remember the layout for fallback use later
            last_floorplan_layout = floorplans[0]

            # Floorplan/Elevation combo builder (your custom module)
            insert_index = generate_floorplan_elevation_slides(
                "templates/floorplan.pptx",
                floorplans,
                pictures,
                prs,
                insert_index,
                event,
                space      # passed correctly
            )

            # Content layout slides using the actual template
            insert_index = generate_layout_content_slides(
                "templates/imageslide.pptx",
                floorplans,
                pictures,
                prs,
                insert_index,
                event
            )
            
            insert_index = generate_floorplan_elevation_slides(
                "templates/Elevation.pptx",
                elevations,
                pictures,
                prs,
                insert_index,
                event,
                space_name=space
            )
            
            template_ppt_inspiration = "templates/inspiration_slides_template1.pptx"

            insert_index = generate_inspiration_slides(
                template_ppt_inspiration,
                inspiration,
                prs,
                insert_index,
                event
            )
            

         if pictures and floorplans and elevations:

            # remember the layout for fallback use later
            last_floorplan_layout = floorplans[0]

            # Floorplan/Elevation combo builder (your custom module)
            insert_index = generate_floorplan_elevation_slides(
                "templates/floorplan.pptx",
                floorplans,
                pictures,
                prs,
                insert_index,
                event,
                space      # passed correctly
            )

            # Content layout slides using the actual template
            insert_index = generate_layout_content_slides(
                "templates/imageslide.pptx",
                floorplans,
                pictures,
                prs,
                insert_index,
                event
            )
            
            insert_index = generate_floorplan_elevation_slides(
                "templates/Elevation.pptx",
                elevations,
                pictures,
                prs,
                insert_index,
                event,
                space_name=space
            )
            
            template_ppt_inspiration = "templates/inspiration_slides_template1.pptx"

            insert_index = generate_inspiration_slides(
                template_ppt_inspiration,
                inspiration,
                prs,
                insert_index,
                event
            )
            


          
         if pictures and floorplans and elevations:

            # remember the layout for fallback use later
            last_floorplan_layout = floorplans[0]

            # Floorplan/Elevation combo builder (your custom module)
            insert_index = generate_floorplan_elevation_slides(
                "templates/floorplan.pptx",
                floorplans,
                pictures,
                prs,
                insert_index,
                event,
                space      # passed correctly
            )

            # Content layout slides using the actual template
            insert_index = generate_layout_content_slides(
                "templates/imageslide.pptx",
                floorplans,
                pictures,
                prs,
                insert_index,
                event
            )
            
            insert_index = generate_floorplan_elevation_slides(
                "templates/Elevation.pptx",
                elevations,
                pictures,
                prs,
                insert_index,
                event,
                space_name=space
            )
            
            template_ppt_inspiration = "templates/inspiration_slides_template1.pptx"

            insert_index = generate_inspiration_slides(
                template_ppt_inspiration,
                inspiration,
                prs,
                insert_index,
                event
            )
            
 
          
          
          
          
          
          
          
                
            
        if pictures and floorplans  and not elevations :

            insert_index = generate_floorplan_elevation_slides(
                "templates/Elevation.pptx",
                elevations,
                pictures,
                prs,
                insert_index,
                event,
                space_name=space
            )    

        # ---------------------------------------------------------
        # 3️⃣ ONLY PICTURES → use previous floorplan layout
        # ---------------------------------------------------------
        if pictures and not floorplans:

            if last_floorplan_layout:
                print(f"✔ Using previous FLOORPLAN layout: {last_floorplan_layout}")
                insert_index = generate_layout_content_slides(
                    "templates/floorplan.pptx",
                    [last_floorplan_layout],
                    pictures,
                    prs,
                    insert_index,
                    event
                )
                insert_index = generate_layout_content_slides(
                    "templates/imageslide.pptx",
                    [last_floorplan_layout],
                    pictures,
                    prs,
                    insert_index,
                    event
                )
            else:
                print("⚠ No previous floorplan found — using first picture fallback.")
                insert_index = generate_layout_content_slides(
                   "templates/imageslide.pptx",
                    [pictures[0]],    # fallback use
                    pictures,
                    prs,
                    insert_index,
                    event
                )

        # ---------------------------------------------------------
        # 4️⃣ INSPIRATION SLIDES
        # ---------------------------------------------------------
        if inspiration:
            template_ppt_inspiration = "templates/inspiration_slides_template1.pptx"

            insert_index = generate_inspiration_slides(
                template_ppt_inspiration,
                inspiration,
                prs,
                insert_index,
                event
            )

    # ---------------------------------------------------------
    # 5️⃣ STYLE SLIDE
    # ---------------------------------------------------------
    if selected_style:
        insert_style_slide(prs, selected_style, insert_index)

    return prs, insert_index




