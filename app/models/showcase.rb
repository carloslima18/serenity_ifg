require 'rubygems'
require 'serenity'

class Showcase
  include Serenity::Generator

  Person = Struct.new(:name, :items)
  Item = Struct.new(:name, :usage)

  def generate_showcase
    @title = 'Serenity inventory'

    mals_items = [Item.new('Moses Brothers Self-Defense Engine Frontier Model B', 'Lock and load')]
    mal = Person.new('Malcolm Reynolds', mals_items)

    jaynes_items = [Item.new('Vera', 'Callahan full-bore auto-lock with a customized trigger, double cartridge and thorough gauge'),
                    Item.new('Lux', 'Ratatata'),
                    Item.new('Knife', 'Cut-throat')]
    jayne = Person.new('Jayne Cobb', jaynes_items)

    @names = ['name_1', 'name_2', 'name_3', 'name_4']

    @crew = [mal, jayne]

    @path_images = []
    @path_images << ::File.join(Rails.public_path, 'dog.png')
    @path_images << ::File.join(Rails.public_path, 'cat.png')

    render_odt "#{Rails.root}/public/files/showcase.docx"
    render_odt "#{Rails.root}/public/files/showcase.odt"
    render_odt "#{Rails.root}/public/files/showcase.ods"
    render_odt "#{Rails.root}/public/files/showcase.xlsx"
  end
end