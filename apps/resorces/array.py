def push(arr, element):
    index = len(arr)
    arr.insert(index, element)
    return arr


def queryset_to_arr(queryset):
    new_list = []
    for v in queryset:
        push(new_list, v)
    return []



def queryset_to_arr_FK(queryset):
    new_list = []
    for v in queryset:
        print(v.id)
    return []