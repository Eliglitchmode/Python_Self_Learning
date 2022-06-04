import os # operating system


def read_file(filename):
    products = []
    with open(filename, 'r', encoding = 'utf-8') as f:
        for line in f:
            if 'Product, Price' in line:
                continue #Jump to next loop and skip the current one 
            name, price = line.strip().split(',')
            products.append([name, price])
    return products


#Two-dimensional list
#Let user input the price & product
def user_input(products):
    while True:
        name = input('Please enter the name of the product: ')
        if name == 'q':  #exit
            break
        price = input('Please enter the price: ')
        price = int(price)
    #   p = []
    #   p.append(name)
    #   p.append(price)
    #   can be written as follows:
    #   p = [name, price]
        products.append([name, price])
    print(products)
    return products

#Print all purchases
def print_products(products):
    for p in products:
        print('The price of ',p[0],'is ', p[1])

#Write into the file
# Use encoding='utf-8' can prevent language not showing properly
def write_file(filename, products):
    with open(filename, 'w', encoding='utf-8') as f:   #with --> help auto-close the file after written
        f.write('Product, Price\n')
        for p in products:
            f.write(p[0] + ',' + str(p[1]) + '\n')


def main():
    filename = 'products.csv'
    if os.path.isfile(filename):
        products = read_file(filename)
        print('yeah! I found it!')
    else: 
        print('Not found....')

    products = user_input(products)
    print_products(products)
    write_file('products.csv', products)

main()
