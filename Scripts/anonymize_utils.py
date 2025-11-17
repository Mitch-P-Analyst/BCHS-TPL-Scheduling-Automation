
def random_name_generator(name_df, number, seed_number=None):
    import random
    """
    name_df : DataFrame with columns ['name', 'gender']
    number  : total number of full names to generate
    
    Returns a list of 'First Last' strings, roughly half female and half male.
    """
    rng = random.Random(seed_number)  # local RNG; doesn't touch global state

    # Filter names by gender
    female_names = name_df.loc[name_df['gender'].str.lower() == 'f', 'name'].tolist()
    male_names   = name_df.loc[name_df['gender'].str.lower() == 'm', 'name'].tolist()

    if not female_names or not male_names:
        raise ValueError("Both female ('f') and male ('m') names are required in name_df.")

    # Decide how many of each gender
    base = number // 2
    extra = number % 2   # 0 or 1

    # Assign extra odd value to male or female
    if extra == 1 and rng.random() < 0.5:
        num_f = base + 1
        num_m = base
    else:
        num_f = base
        num_m = base + extra

    output_list = []

    # Generate female names
    for _ in range(num_f):
        first_name = rng.choice(female_names)
        last_name  = rng.choice(female_names)
        output_list.append(f"{first_name} {last_name}")

    # Generate male names
    for _ in range(num_m):
        first_name = rng.choice(male_names)
        last_name  = rng.choice(male_names)
        output_list.append(f"{first_name} {last_name}")

    # Shuffle so they’re not grouped by gender
    rng.shuffle(output_list)

    return output_list





def first_name_generator(name_df, number, seed_number=None):
    import random
    """
    name_df : DataFrame with columns ['name', 'gender']
    number  : total number of first names to generate
    seed_number : optional; if given, makes output reproducible
    """

    rng = random.Random(seed_number)  # local RNG; doesn't touch global state

    # Filter names by gender
    female_names = name_df.loc[name_df['gender'].str.lower() == 'f', 'name'].tolist()
    male_names   = name_df.loc[name_df['gender'].str.lower() == 'm', 'name'].tolist()

    if not female_names or not male_names:
        raise ValueError("Both female ('f') and male ('m') names are required in name_df.")

    # Decide how many of each gender
    base = number // 2
    extra = number % 2  # 0 or 1

    # Assign extra odd value to male or female
    if extra == 1 and rng.random() < 0.5:
        num_f = base + 1
        num_m = base
    else:
        num_f = base
        num_m = base + extra

    output_list = []

    # Generate female names
    for _ in range(num_f):
        first_name = rng.choice(female_names)
        output_list.append(first_name)

    # Generate male names
    for _ in range(num_m):
        first_name = rng.choice(male_names)
        output_list.append(first_name)

    rng.shuffle(output_list)

    return output_list

